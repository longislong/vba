Option Explicit

'[参照設定]
'Microsoft HTML Object Library
'Microsoft Internet Controls

'◆WindowsAPIの引っ張り込み
Private Declare PtrSafe Function PostMessage Lib "user32.dll" Alias "PostMessageA" _
            (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
            (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
            ByVal lpszClass As String, ByVal lpszWindow As String) As Long

Private Declare PtrSafe Function SetForegroundWindow Lib "user32.dll" _
            (ByVal hWnd As Long) As Long

