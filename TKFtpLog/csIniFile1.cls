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
'物件設計師：徐宇璇   Mario
'物件版本  ：1.00    1999.05.26.
'            1.01    1999.09.07.  修正讀寫中文資料錯誤，將Len改為LenB
'物件名稱  ：vbIniFile
'屬性      ：
'           IniFileName    設定或傳回欲讀寫的ini檔名
'方法      ：
'           WriteData      寫入
'           ReadData       讀取


'讀寫INI File 所需要的API宣告
'宣告程序私有成員函數
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'私有常數宣告
Private Const IniExtName = ".ini"

'私有變數
Private mvarIniFileName As String   'Ini的檔案名稱

'私有成員函數
'檢查是否有副檔名
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

'檢查是否有指定讀寫路徑
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

'增加目前執行檔的路徑
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

