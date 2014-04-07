Attribute VB_Name = "modAddInRegistration"
Option Explicit

'Code from:
'http://support.microsoft.com/support/kb/articles/Q238/2/28.ASP
'HOWTO: Build an Office 2000 COM Add-In in Visual Basic

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As _
Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" _
Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
ByVal cbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) _
As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS = &H1F0037
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

'These are the settings for your Add-in...
Private Const PROGID As String = "OutlookQuote.clsAddIn"
Private Const DESCRIPTION As String = "Add a quote in your mails"
Private Const LOADBEHAVIOR As Long = 3
Private Const SAFEFORCOMMANDLINE As Long = 0

Public Sub RegisterAll()
   RegisterOfficeAddin "Access"
   RegisterOfficeAddin "Excel"
   RegisterOfficeAddin "FrontPage"
   RegisterOfficeAddin "Outlook"
   RegisterOfficeAddin "PowerPoint"
   RegisterOfficeAddin "Word"
End Sub

Public Sub UnregisterAll()
   UnRegisterOfficeAddin "Access"
   UnRegisterOfficeAddin "Excel"
   UnRegisterOfficeAddin "FrontPage"
   UnRegisterOfficeAddin "Outlook"
   UnRegisterOfficeAddin "PowerPoint"
   UnRegisterOfficeAddin "Word"
End Sub

Public Sub RegisterOfficeAddin(sTargetApp As String)
   Dim sRegKey As String
   Dim nRet As Long, dwTmp As Long
   Dim hKey As Long

   sRegKey = "Software\Microsoft\Office\" & sTargetApp _
      & "\Addins\" & PROGID

   nRet = RegCreateKeyEx(HKEY_CURRENT_USER, sRegKey, 0, _
      vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, dwTmp)
   
   If nRet = 0 Then
      If dwTmp = REG_CREATED_NEW_KEY Then
         Call RegSetValueEx(hKey, "FriendlyName", 0, _
            REG_SZ, ByVal PROGID, Len(PROGID))
         Call RegSetValueEx(hKey, "Description", 0, _
            REG_SZ, ByVal DESCRIPTION, Len(DESCRIPTION))
         Call RegSetValueEx(hKey, "LoadBehavior", 0, _
            REG_DWORD, LOADBEHAVIOR, 4)
         Call RegSetValueEx(hKey, "CommandLineSafe", 0, _
            REG_DWORD, SAFEFORCOMMANDLINE, 4)
      End If
      Call RegCloseKey(hKey)
   End If

End Sub

Public Sub UnRegisterOfficeAddin(sTargetApp As String)
   Dim sRegKey As String
   sRegKey = "Software\Microsoft\Office\" & sTargetApp _
      & "\Addins\" & PROGID

    Call RegDeleteKey(HKEY_CURRENT_USER, sRegKey)

End Sub

