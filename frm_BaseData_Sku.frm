VERSION 5.00
Begin VB.Form frm_BaseData_Sku 
   Caption         =   "�f���򥻸�ƺ��@"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6510
   Begin VB.CommandButton cmd_Only_sku 
      BackColor       =   &H000000FF&
      Caption         =   "�u��s�~�W"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox text_bestdescr 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmd_Exit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "��  �}"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Index           =   1
      Left            =   5040
      Picture         =   "frm_BaseData_Sku.frx":0000
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      Top             =   3720
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Tab0_Save 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�s  ��"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3990
      Picture         =   "frm_BaseData_Sku.frx":0442
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      Top             =   3720
      Width           =   1065
   End
   Begin VB.TextBox txt_CaseWT 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txt_H 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txt_W 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txt_L 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox text_sku 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   960
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '�z��
      Caption         =   "�ըƹF�~�W�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  '���z��
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      Height          =   1080
      Index           =   1
      Left            =   3840
      Top             =   3600
      Width           =   2400
   End
   Begin VB.Label Lab_WT 
      BackStyle       =   0  '�z��
      Caption         =   "LABT01"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '�z��
      Caption         =   "�p��쭫�q : "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1545
   End
   Begin VB.Label Lab_CBF 
      BackStyle       =   0  '�z��
      Caption         =   "LABT01"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '�z��
      Caption         =   "�p���~�n : "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1545
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '�z��
      Caption         =   "�C�c���q�G               (����)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   3345
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '�z��
      Caption         =   "�C�c�W��G��              �e              ��              (����)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   5985
   End
   Begin VB.Label Lab_Casecnt 
      BackStyle       =   0  '�z��
      Caption         =   "LABT01"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "�J�ơG"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Lab_Descr 
      BackStyle       =   0  '�z��
      Caption         =   "LABT01"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  '�z��
      Caption         =   "�f���G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   810
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�f�D�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '�z��
      Caption         =   "�~�W�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Lab_storer 
      BackStyle       =   0  '�z��
      Caption         =   "LABT01���ӨȰ��ѥ��������q"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   870
      TabIndex        =   8
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frm_BaseData_Sku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Exit_Click(Index As Integer)
    Unload Me '�������{��
End Sub



Private Sub cmd_Tab0_Save_Click()
Dim i As Integer
    If Len(Trim(text_sku.Text)) = 0 Then
        msg_text = "�Х��T�{�f��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(Lab_WT.Caption)) = 0 Then
        msg_text = "���n�L���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Len(Trim(Lab_CBF.Caption)) = 0 Then
        msg_text = "���q�L���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Val(Trim(Lab_WT.Caption)) = 0 Then
        msg_text = "���n���i���s"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    If Val(Trim(Lab_CBF.Caption)) = 0 Then
        msg_text = "���q���i���s"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    On Error GoTo err_handle
    If Len(Trim(txt_L.Text)) > 0 And Len(Trim(txt_W.Text)) > 0 And Len(Trim(txt_H.Text)) > 0 Then
        
        cmd_Only_sku_Click  '��s�~�W
        Tran_Level = cn.BeginTrans
        '��s���n���q
        str_SQL = "Update Exceed_ABT..sku set STDGROSSWGT='" & Trim(Lab_WT.Caption) & "',STDCUBE='" & Trim(Lab_CBF.Caption) & "' where sku = '" & Trim(text_sku.Text) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "Update Exceed_ABT..pack set Exceed_ABT..pack.LengthUOM1='" & Trim(txt_L.Text) & "',Exceed_ABT..pack.WidthUOM1='" & Trim(txt_W.Text) & "',Exceed_ABT..pack.HeightUOM1='" & Trim(txt_H.Text) & "' from Exceed_ABT..pack join Exceed_ABT..sku on Exceed_ABT..sku.packkey = Exceed_ABT..pack.packkey and Exceed_ABT..sku.sku = '" & Trim(text_sku.Text) & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        cn.CommitTrans: Tran_Level = 0
    Else
        msg_text = "���e����Ƥ�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
Exit Sub
    
err_handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans: Tran_Level = 0
   End If
   Dim tmpString As String
   Screen.MousePointer = vbDefault
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "--�s��", Me.Caption, "cmd_Tab0_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    
End Sub

Private Sub cmd_Only_sku_Click()
    '�L�otab,�ť�
    cmd_Only_sku.Enabled = False: cmd_Tab0_Save.Enabled = False
    text_sku.Text = myExCharFilter(text_sku.Text)
    text_bestdescr.Text = myExCharFilter(text_bestdescr.Text)
    Lab_WT.Caption = myExCharFilter(Lab_WT.Caption)
    Lab_CBF.Caption = myExCharFilter(Lab_CBF.Caption)
    txt_L.Text = myExCharFilter(txt_L.Text)
    txt_W.Text = myExCharFilter(txt_W.Text)
    txt_H.Text = myExCharFilter(txt_H.Text)
        Tran_Level = cn.BeginTrans
        If Len(Trim(text_bestdescr.Text)) = 0 Then '����s�ըƹF�~�W
        Else
            '�P�_�O�_�w�g���ʨƹF�~�W , ���h��s,�S���h�s�W
                '��s�ըƹF�~�W
                   str_SQL = "update storersku set descr = '" & text_bestdescr.Text & "' where sku = '" & Trim(text_sku.Text) & "' and storerkey = 'LABT01'"
                   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
             If RowsAffect = 0 Then        '�L��ƥi��s�h�s�W
                '�s�W�ըƹF�~�W
                   str_SQL = "insert into storersku(storerkey,sku,storersku,descr)" & _
                             "values('LABT01','" & Trim(text_sku.Text) & "','" & Trim(text_sku.Text) & "','" & text_bestdescr.Text & "')"
                   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
             End If
       End If
        cn.CommitTrans: Tran_Level = 0
        msg_text = "��s���\"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title

End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�f���򥻸�ƺ��@"
End Sub

Private Sub Form_Load()
    '�]�w Form �j�p�B��m
    Me.Height = 5325: Me.Width = 6450
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    'Screen.MousePointer = vbDefault
    Lab_Casecnt.Caption = ""
    Lab_WT.Caption = ""
    Lab_CBF.Caption = ""
End Sub

Private Sub text_sku_Change()
cmd_Tab0_Save.Enabled = False
End Sub

Private Sub text_sku_KeyPress(KeyAscii As Integer)
    On Error GoTo err_handle
    If KeyAscii = 13 Then
        If Len(Trim(text_sku.Text)) = 0 Then
            Screen.MousePointer = vbDefault
            msg_text = "�d�ߤ��i���ŭȡI"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Exit Sub
        Else
        
            str_sku = myExCharFilter(text_sku.Text)
            
            str_SQL = "select s.sku,s.descr ,isnull(ss.descr,' ') as bestdescr,p.casecnt,p.LengthUOM1,p.WidthUOM1,p.HeightUOM1,s.STDGROSSWGT,s.STDCUBE " & _
                      "from Exceed_ABT..sku s inner join Exceed_ABT..pack p on s.packkey=p.packkey left join storersku ss on s.sku = ss.sku and s.storerkey = ss.storerkey " & _
                      "where s.storerkey='LABT01' and s.sku='" & str_sku & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            
            If Not tmp_Rs.EOF Then
            '������ƪ���
                Lab_Descr.Caption = tmp_Rs.Fields("descr").Value
                Lab_Casecnt.Caption = tmp_Rs.Fields("casecnt").Value
                Lab_CBF.Caption = tmp_Rs.Fields("STDCUBE").Value
                Lab_WT.Caption = tmp_Rs.Fields("STDGROSSWGT").Value
                text_bestdescr.Text = Trim(tmp_Rs.Fields("bestdescr").Value)
                txt_L.Text = tmp_Rs.Fields("LengthUOM1").Value
                txt_W.Text = tmp_Rs.Fields("WidthUOM1").Value
                txt_H.Text = tmp_Rs.Fields("HeightUOM1").Value
                txt_CaseWT = ""
                cmd_Tab0_Save.Enabled = True
                cmd_Only_sku.Enabled = True
                
            Else
                Screen.MousePointer = vbDefault
                msg_text = "�~���G" & Trim(str_sku) & " �L��ơA�нT�{�C"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                Lab_Descr.Caption = ""
                Lab_Casecnt.Caption = ""
                Lab_CBF.Caption = ""
                Lab_WT.Caption = ""
                txt_L.Text = ""
                txt_W.Text = ""
                txt_H.Text = ""
            End If
        End If
    End If
Exit Sub
    
err_handle:
   Dim tmpString As String
   Screen.MousePointer = vbDefault
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "--", Me.Caption, "text_sku_KeyPress", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub txt_CaseWT_Change()
    'Lab_WT
    If Val(txt_CaseWT) = 0 Then Exit Sub
    If Len(Lab_Casecnt.Caption) > 0 Then
        If (Lab_Casecnt.Caption) = 0 Then
            Lab_WT.Caption = Trim(txt_CaseWT.Text)
        Else
            Lab_WT.Caption = Val(txt_CaseWT.Text) / (Lab_Casecnt.Caption)
        End If
    Lab_WT.Caption = Round(Lab_WT.Caption, 10)
    End If
End Sub

Private Sub txt_CaseWT_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 45 Or KeyAscii > 58) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_H_Change()
    If Len(Trim(txt_L.Text)) > 0 And Len(Trim(txt_W.Text)) > 0 And Len(Trim(txt_H.Text)) > 0 Then
        If Len(Lab_Casecnt.Caption) > 0 Then
            If (Lab_Casecnt.Caption) = 0 Then
                Lab_CBF.Caption = ((txt_L.Text) * (txt_W.Text) * (txt_H.Text) * 0.0000353)
            Else
                Lab_CBF.Caption = ((txt_L.Text) * (txt_W.Text) * (txt_H.Text) * 0.0000353) / (Lab_Casecnt.Caption)
            End If
        End If
    Lab_CBF.Caption = Round(Lab_CBF.Caption, 10)
    End If
End Sub

Private Sub txt_H_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 45 Or KeyAscii > 58) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_L_Change()
    If Len(Trim(txt_L.Text)) > 0 And Len(Trim(txt_W.Text)) > 0 And Len(Trim(txt_H.Text)) > 0 Then
        If Len(Lab_Casecnt.Caption) > 0 Then
            If (Lab_Casecnt.Caption) = 0 Then
                Lab_CBF.Caption = ((txt_L.Text) * (txt_W.Text) * (txt_H.Text) * 0.0000353)
            Else
                Lab_CBF.Caption = ((txt_L.Text) * (txt_W.Text) * (txt_H.Text) * 0.0000353) / (Lab_Casecnt.Caption)
            End If
        End If
    Lab_CBF.Caption = Round(Lab_CBF.Caption, 10)
    End If
End Sub

Private Sub txt_L_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 45 Or KeyAscii > 58) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_W_Change()
    If Len(Trim(txt_L.Text)) > 0 And Len(Trim(txt_W.Text)) > 0 And Len(Trim(txt_H.Text)) > 0 Then
        If Len(Lab_Casecnt.Caption) > 0 Then
            If (Lab_Casecnt.Caption) = 0 Then
                Lab_CBF.Caption = ((txt_L.Text) * (txt_W.Text) * (txt_H.Text) * 0.0000353)
            Else
                Lab_CBF.Caption = ((txt_L.Text) * (txt_W.Text) * (txt_H.Text) * 0.0000353) / (Lab_Casecnt.Caption)
            End If
        End If
    Lab_CBF.Caption = Round(Lab_CBF.Caption, 10)
    End If
End Sub

Private Sub txt_W_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 45 Or KeyAscii > 58) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
