Attribute VB_Name = "MCIWndBase"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function MCIWndRegisterClassAPI Lib "msvfw32" Alias "MCIWndRegisterClass" () As Long
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long
#Else
Private Declare Function MCIWndRegisterClassAPI Lib "msvfw32" Alias "MCIWndRegisterClass" () As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
#End If
Private MCIWndRefCount As Long

Public Sub MCIWndRegisterClass()
If MCIWndRefCount = 0 Then MCIWndRegisterClassAPI
MCIWndRefCount = MCIWndRefCount + 1
End Sub

Public Sub MCIWndReleaseClass()
MCIWndRefCount = MCIWndRefCount - 1
If MCIWndRefCount = 0 Then UnregisterClass StrPtr("MCIWndClass"), App.hInstance
End Sub
