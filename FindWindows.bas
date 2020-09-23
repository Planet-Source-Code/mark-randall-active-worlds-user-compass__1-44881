Attribute VB_Name = "FindWindows"
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function GetWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal wCmd As Long) As Long

Public Declare Function GetWindowText Lib "user32" _
   Alias "GetWindowTextA" _
  (ByVal hWnd As Long, _
   ByVal lpString As String, _
   ByVal cch As Long) As Long

Public Declare Function GetClassName Lib "user32" _
   Alias "GetClassNameA" _
  (ByVal hWnd As Long, _
   ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5

