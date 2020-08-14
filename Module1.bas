Attribute VB_Name = "Module1"
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_LWIN = &H5B
Public Const KEYEVENTF_KEYUP = &H2


