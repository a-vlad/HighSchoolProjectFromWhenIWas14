Attribute VB_Name = "MouseClip"
 Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Public Function ClipYes(lpRect As RECT)
ClipCursor lpRect
End Function

Public Function ReleaseYes()
ClipCursor ByVal vbNullString
End Function


Public Function FillRect(frm As Object, lpRect As RECT)
GetWindowRect frm.hwnd, lpRect
End Function

