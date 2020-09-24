Attribute VB_Name = "mAPI"
Option Explicit
DefLng A-Z 'we're 32 bit

'Splash duration
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds)

Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long

'About box
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Send Mail, Open RTF file
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL          As Long = 1
Public Const SE_NO_ERROR            As Long = 33 'Values below 33 are error returns

'VB Font properties
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey, ByVal lpSubKey As String, ByVal ulOptions, ByVal samDesired, phkResult)
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey, ByVal lpValueName As String, ByVal lpReserved, lpType, lpData As Any, lpcbData)
Public Declare Sub RegCloseKey Lib "advapi32.dll" (ByVal hKey)

Public Const HKEY_CURRENT_USER      As Long = &H80000001
Public Const KEY_QUERY_VALUE        As Long = 1
Public Const REG_OPTION_RESERVED    As Long = 0
Public Const ERROR_NONE             As Long = 0
Public Const WM_PASTE               As Long = &H302

'Some common Keys for Imagelist in fXRef
Public Const KeyFolderOpenImg       As String = "FilO"
Public Const KeyFolderClosedImg     As String = "FilC"
Public Const KeyReferenceImg        As String = "Ref"

Public VisibleNode                  As MSComctlLib.Node

Public UnrefFound                   As Boolean
Public DuplFound                    As Boolean

Public CompoFound                   As Boolean
Public LastWindowState
Public LastFoundIndex
Public WholeWord

Public LastSrchFor                  As String
Public SourceDir                    As String

Public Sub SendMeMail(FromhWnd, Subject As String)

    If ShellExecute(FromhWnd, vbNullString, "mailto:UMGEDV@AOL.COM?subject=" & Subject & " &body=Hi Ulli,", vbNullString, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
    End If

End Sub

Public Function AppDetails() As String

    With App
        AppDetails = .ProductName & " Version " & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Function

Public Sub Inc(What As Long, Optional By As Long = 1)

    What = What + By

End Sub

':) Ulli's VB Code Formatter V2.4.4 (21.10.2001 11:53:57) 45 + 24 = 69 Lines
