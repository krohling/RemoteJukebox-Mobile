Attribute VB_Name = "DisableX"
Private Const MF_BYPOSITION = &H400&
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
                  ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
                  ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" _
                 (ByVal hMenu As Long) As Long


Public Sub DisableX(lhwnd As Long)
Dim lSysMenu As Long
Dim lItemCount As Long
Dim lRet As Long
lSysMenu = GetSystemMenu(lhwnd, False)
lItemCount = GetMenuItemCount(lSysMenu)
lRet = RemoveMenu(lSysMenu, lItemCount - 1, MF_BYPOSITION)
lRet = RemoveMenu(lSysMenu, lItemCount - 2, MF_BYPOSITION)
lRet = RemoveMenu(lSysMenu, lItemCount - 3, MF_BYPOSITION)
lRet = RemoveMenu(lSysMenu, lItemCount - 4, MF_BYPOSITION)
End Sub


