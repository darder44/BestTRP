VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'����]�p�v�G�}�t�   Mario
'���󪩥�  �G1.00    1999.05.26.
'            1.01    1999.09.07.  �ץ�Ū�g�����ƿ��~�A�NLen�אּLenB
'����W��  �GvbIniFile
'�ݩ�      �G
'           IniFileName    �]�w�ζǦ^��Ū�g��ini�ɦW
'��k      �G
'           WriteData      �g�J
'           ReadData       Ū��


'Ū�gINI File �һݭn��API�ŧi
'�ŧi�{�Ǩp���������
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'�p���`�ƫŧi
Private Const IniExtName = ".ini"

'�p���ܼ�
Private mvarIniFileName As String   'Ini���ɮצW��

'�p���������
'�ˬd�O�_�����ɦW
Private Function CheckIniExtionName(ByVal vIniFileName As String) As Boolean
   Dim vIniExtLen As Long
   vIniExtLen = LenB(vIniFileName)
   Dim i As Long
   
   For i = 1 To vIniExtLen
       If Mid(vIniFileName, i, 1) = "." And _
          LenB(Mid(vIniFileName, i, vIniExtLen - i)) > 0 Then
          CheckIniExtionName = True
          Exit Function
       End If
   Next i
   
   CheckIniExtionName = False
End Function

'�ˬd�O�_�����wŪ�g���|
Private Function CheckIniFilePath(ByVal vIniFileName As String) As Boolean
   Dim vIniPathLen As Long
   vIniPathLen = LenB(vIniFileName)
   Dim i As Long
   
   For i = 1 To vIniPathLen
       If Mid(vIniFileName, i, 1) = "\" And _
          LenB(Mid(vIniFileName, i, vIniPathLen - i)) > 0 Then
          CheckIniFilePath = True
          Exit Function
       End If
   Next i
   
   CheckIniFilePath = False
End Function

'�W�[�ثe�����ɪ����|
Private Function AppPath() As String
  If Right(App.Path, 1) = "\" Then
     AppPath = App.Path
  Else
     AppPath = App.Path & "\"
  End If
End Function

Public Property Let FileName(ByVal vData As String)
    If LenB(vData) = 0 Then
       mvarIniFileName = AppPath & App.EXEName & IniExtName
    Else
       If CheckIniExtionName(vData) Then
          mvarIniFileName = vData
       Else
          If Right(vData, 1) = "." Then
             mvarIniFileName = vData & "ini"
          Else
             mvarIniFileName = vData & ".ini"
          End If
       End If
       
       If Not CheckIniFilePath(vData) Then
          mvarIniFileName = AppPath & mvarIniFileName
       End If
       
    End If
End Property

Public Property Get FileName() As String
    iniFileName = mvarIniFileName
End Property

Public Function ReadData(ByVal SectionName As String, ByVal keyName As String, DefaultValue As Variant) As Variant
    Dim rString As String
    rString = String(1024, 0)
    Dim rStrLen As Long
    
    rStrLen = GetPrivateProfileString(SectionName, keyName, CStr(DefaultValue), rString, LenB(rString), mvarIniFileName)
    rString = Left(rString, rStrLen)
    
    ReadData = rString
End Function

Public Function WriteData(ByVal SectionName As String, ByVal keyName As String, KeyValue As Variant) As Boolean
   Dim wString As String
   wString = CStr(KeyValue)
   
   Dim ret As Long
   ret = WritePrivateProfileString(SectionName, keyName, wString, mvarIniFileName)
   If ret <> 0 Then
      WriteData = True
   Else
      WriteDate = False
   End If
   
End Function

Private Sub Class_Initialize()
  mvarIniFileName = App.EXEName & ".ini"
End Sub
