VERSION 5.00
Begin VB.Form frm_WaitWindows 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�t�Ϊ��A�E�E�E�E�E�E"
   ClientHeight    =   2610
   ClientLeft      =   3420
   ClientTop       =   3315
   ClientWidth     =   6180
   Icon            =   "frm_WaitWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_Cancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   60
      Picture         =   "frm_WaitWindows.frx":030A
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5610
      Top             =   150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   4380
      Picture         =   "frm_WaitWindows.frx":0BD4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   285
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   1575
      Picture         =   "frm_WaitWindows.frx":0EDE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      TabIndex        =   4
      Top             =   1935
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      TabIndex        =   3
      Top             =   1590
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Top             =   1230
      Width           =   600
   End
End
Attribute VB_Name = "frm_WaitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dtStart As Date
Dim updLabel3 As Boolean

Private Sub cmd_Cancel_Click()
Select Case UCase(Me.Tag)
       Case "TRANSFERTOEXCEL"
            fgTransferToExcel = False
End Select

End Sub

Private Sub Form_Activate()
dtStart = Now: updLabel3 = True
Select Case UCase(Me.Tag)
       Case "FRM_MDIFORM"
            Label1.Caption = "�P��Ʈw�إ߳s�u���A�еy��E�E�E�E�E�E"
            Label2.Caption = "�s�u�_�l�ɶ��G" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "�s�u�ɶ��G 00 �� 00 ��"
            Call DB_connect
            updLabel3 = False
            Unload Me
       Case "IMPORTDATA"       '��ƶפJ�@�~
            Label1.Caption = "��ƶפJ�@�~���椤�A�еy��E�E�E�E�E�E"
            Label2.Caption = "�_�l�ɶ��G" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "�פJ�ɶ��G 00 �� 00 ��"
       Case "TRANSFERTOEXCEL"  '���Ʀ� Excel
            Label1.Caption = "[��s Excel ��] �@�~���椤�A�еy��E�E�E�E�E�E"
            Label2.Caption = "�_�l�ɶ��G" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "���ɮɶ��G 00 �� 00 ��"
            cmd_Cancel.Visible = True
       Case Else
            Label1.Caption = "�t�Χ@�~���椤�A�еy��E�E�E�E�E�E"
            Label2.Caption = "�_�l�ɶ��G" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "�d�߮ɶ��G 00 �� 00 ��"
End Select
End Sub

Private Sub Form_Load()
Me.Height = 3000: Me.Width = 6300
Me.Left = ((frm_MDIForm.ScaleWidth - Me.Width) / 2) + 600
Me.Top = ((frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_WaitWindows = Nothing
End Sub

Private Sub Timer1_Timer()
Select Case UCase(Me.Tag)
       Case "FRM_MDIFORM"
            If updLabel3 = True Then
               Label3.Caption = "�s�u�ɶ��G  " & Format(Now - dtStart, "nn") & " �� " & Format(Now - dtStart, "ss") & " ��"
            End If
       Case "TRANSFERTOEXCEL"
            If updLabel3 = True Then
               Label3.Caption = "���ɮɶ��G  " & Format(Now - dtStart, "nn") & " �� " & Format(Now - dtStart, "ss") & " ��"
            End If
       Case "IMPORTDATA"
            If updLabel3 = True Then
               Label3.Caption = "�פJ�ɶ��G  " & Format(Now - dtStart, "nn") & " �� " & Format(Now - dtStart, "ss") & " ��"
            End If
      Case Else
            If updLabel3 = True Then
               Label3.Caption = "����ɶ��G  " & Format(Now - dtStart, "nn") & " �� " & Format(Now - dtStart, "ss") & " ��"
            End If
End Select
End Sub
