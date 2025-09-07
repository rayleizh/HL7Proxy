Attribute VB_Name = "modRegEnum"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
     lpcName As Long, lpReserved As Long, ByVal lpClass As String, _
     lpcClass As Long, lpftLastWriteTime As Any) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const KEY_READ As Long = &H20019
Private Const KEY_WOW64_32KEY As Long = &H200

Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&

Public Function EnumMeridianServiceKeys() As Collection
    Dim result As New Collection
    Dim hKey As Long, rc As Long
    
    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Meridian", 0&, KEY_READ Or KEY_WOW64_32KEY, hKey)
    If rc <> ERROR_SUCCESS Then
        Set EnumMeridianServiceKeys = result
        Exit Function
    End If
    
    Dim idx As Long: idx = 0
    Do
        Dim nameBuf As String * 256
        Dim nameLen As Long: nameLen = 255
        rc = RegEnumKeyEx(hKey, idx, nameBuf, nameLen, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&)
        If rc = ERROR_NO_MORE_ITEMS Then Exit Do
        If rc = ERROR_SUCCESS Then
            Dim k As String: k = Left$(nameBuf, nameLen)
            If Left$(k, 9) = "HL7Server" Then result.Add k
        End If
        idx = idx + 1
    Loop
    
    RegCloseKey hKey
    Set EnumMeridianServiceKeys = result
End Function
