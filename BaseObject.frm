VERSION 5.00
Begin VB.Form BaseObject 
   Caption         =   "BaseObject"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   11280
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame Frame7 
      Caption         =   "�\��"
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
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   9615
      Begin VB.CommandButton cmdPreView 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�w��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   2880
         Picture         =   "BaseObject.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�C�L"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   4200
         Picture         =   "BaseObject.frx":1708A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   13
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��Excel"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   5520
         Picture         =   "BaseObject.frx":17394
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   240
         Picture         =   "BaseObject.frx":1868E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���]"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   1560
         Picture         =   "BaseObject.frx":18998
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�M��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   6840
         Picture         =   "BaseObject.frx":18CAA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�T�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   8160
         Picture         =   "BaseObject.frx":1A9A4
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
         Top             =   360
         Width           =   1185
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000C&
         BackStyle       =   1  '���z��
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         Height          =   1485
         Left            =   120
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�\��"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFF80&
         Caption         =   "�s��"
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
         Height          =   1125
         Left            =   4200
         Picture         =   "BaseObject.frx":1C69E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�R��"
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
         Height          =   1125
         Left            =   2880
         Picture         =   "BaseObject.frx":1C9A8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�ק�"
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
         Height          =   1125
         Left            =   1560
         Picture         =   "BaseObject.frx":1D9EA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdAddNew 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�s�W"
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
         Height          =   1125
         Left            =   240
         Picture         =   "BaseObject.frx":2423C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF80FF&
         Caption         =   "���}"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   6840
         Picture         =   "BaseObject.frx":26366
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   5520
         Picture         =   "BaseObject.frx":4FF78
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000C&
         BackStyle       =   1  '���z��
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         Height          =   1365
         Left            =   120
         Top             =   240
         Width           =   8055
      End
   End
End
Attribute VB_Name = "BaseObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
