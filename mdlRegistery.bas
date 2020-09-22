Attribute VB_Name = "mdlRegistery"
' =========================================
'  Model Registery Keys/Strings Create/Delete/List v1.3
'  by DarkJedi
'  DarkJedi@hotmail.co.il
' ==========================================

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal HKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal HKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
(ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName _
As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, _
lpftLastWriteTime As Any) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
(ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As _
Byte, lpcbData As Long) As Long

' -------------------------------------------
' Function name: GetKeys
' -------------------------------------------
' Syntax: GetKeys HKey, Path, ListBoxName
' Exapmle: GetKeys HKEY_LOCAL_MACHINE, "Software\NewKey", Form1.List1
' -------------------------------------------

Function GetKeys(ByVal H_Key As Long, ByVal HSubDir As String, lstList As ListBox)
    Dim HKey As Long, Counter As Long, sSave As String
 
    RegOpenKey H_Key, HSubDir, HKey
    Do
        sSave = String(255, 0)
        If RegEnumKeyEx(HKey, Counter, sSave, 255, 0, vbNullString, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
        lstList.AddItem StripTerminator(sSave)
        Counter = Counter + 1
    Loop
    RegCloseKey HKey
End Function

' -------------------------------------------
' Sub name: CreateKey
' -------------------------------------------
' Syntax: CreateKey hKey, srtPath
' Exapmle: CreateKey HKEY_LOCAL_MACHINE, "Software\NewKey"
' -------------------------------------------

Public Sub CreateKey(HKey As Long, strPath As String)
    Dim keyhand&
    RegCreateKey HKey, strPath, keyhand&
    RegCloseKey keyhand&
End Sub

' -------------------------------------------
' Sub name: DeleteKey
' -------------------------------------------
' Syntax: DeleteKey hKey, srtPath
' Exapmle: DeleteKey HKEY_LOCAL_MACHINE, "Software\NewKey"
' -------------------------------------------

Public Sub DeleteKey(HKey As Long, ByVal strPath As String)
    RegDeleteKey HKey, strPath
End Sub

' -------------------------------------------
' Function name: GetValues
' -------------------------------------------
' Syntax: GetValues HKey, Path, ListBoxName
' Exapmle: GetValues HKEY_LOCAL_MACHINE, "Software\NewKey", Form1.List1
' -------------------------------------------

Function GetValues(ByVal H_Key As String, ByVal HSubKey As String, lstList As ListBox)
 Dim HKey As Long, Counter As Long, sSave As String
  RegOpenKey H_Key, HSubKey, HKey
  Counter = 0
  Do
     sSave = String(255, 0)
     If RegEnumValue(HKey, Counter, sSave, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
     getS = StripTerminator(sSave)
     If getS = "" Then getS = "(default)"
     lstList.AddItem getS
     Counter = Counter + 1
  Loop
  RegCloseKey HKey
End Function

' -------------------------------------------
' Sub name: DeleteSettingString
' -------------------------------------------
' Syntax: DeleteSettingString hKey, srtPath, strValue
' Exapmle: DeleteSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft", "TestString"
' -------------------------------------------

Public Sub DeleteSettingString(ByVal HKey As Long, _
ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegOpenKey(HKey, strPath, hCurKey)

lRegResult = RegDeleteValue(hCurKey, strValue)

lRegResult = RegCloseKey(hCurKey)

End Sub

' -------------------------------------------
' Sub name: SaveSettingString
' -------------------------------------------
' Syntax: SaveSettingString hKey, srtPath
' Exapmle: SaveSettingString HKEY_LOCAL_MACHINE, "Software\NewKey"
' -------------------------------------------

Public Sub SaveSettingString(HKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal strData As String)
    Dim keyhand&
    RegCreateKey HKey, strPath, keyhand&
    RegSetValueEx keyhand&, strValue, 0, 1, ByVal strData, Len(strData)
    RegCloseKey keyhand&
End Sub

' -------------------------------------------
' Sub name: GetSettingString
' -------------------------------------------
' Syntax: GetSettingString hKey, srtPath, strValue
' Exapmle: Value1 = GetSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft", "TestString"
' -------------------------------------------

Public Function GetSettingString(ByVal HKey As Long, _
ByVal strPath As String, ByVal strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

' Set up default value
If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(HKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If

Else
' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

'--------------------------------------
'Private Function for String Termination
'Input: Softwareeeeeeeeeeeeeeeeee...
'Output: Software
'--------------------------------------
Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    'Search the first chr$(0)
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

