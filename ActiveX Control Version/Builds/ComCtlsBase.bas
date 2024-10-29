Attribute VB_Name = "ComCtlsBase"
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If

#Const ImplementPreTranslateMsg = (VBCCR_OCX <> 0)

Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As LongPtr
dwHoverTime As Long
End Type
Private Type TMSG
hWnd As LongPtr
Message As Long
wParam As LongPtr
lParam As LongPtr
Time As Long
PT As POINTAPI
End Type
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type TLOCALESIGNATURE
lsUsb(0 To 15) As Byte
lsCsbDefault(0 To 1) As Long
lsCsbSupported(0 To 1) As Long
End Type
Private Type TOOLINFO
cbSize As Long
uFlags As Long
hWnd As LongPtr
uId As LongPtr
RC As RECT
hInst As LongPtr
lpszText As LongPtr
lParam As LongPtr
End Type
#If VBA7 Then
Public Declare PtrSafe Function ComCtlsObjAddRef Lib "msvbvm60.dll" Alias "__vbaObjAddref" (ByVal lpObject As LongPtr) As Long
Public Declare PtrSafe Function ComCtlsObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (ByRef Destination As Any, ByVal lpObject As LongPtr) As Long
Public Declare PtrSafe Function ComCtlsObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As LongPtr) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare PtrSafe Function MCIWndRegisterClass Lib "msvfw32" () As Long
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long
Private Declare PtrSafe Function GetClassLong Lib "user32" Alias "GetClassLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As TMSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadID As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetKeyboardLayout Lib "user32" (ByVal dwThreadID As Long) As LongPtr
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As LongPtr
Private Declare PtrSafe Function ImmIsIME Lib "imm32" (ByVal hKL As LongPtr) As Long
Private Declare PtrSafe Function ImmCreateContext Lib "imm32" () As LongPtr
Private Declare PtrSafe Function ImmDestroyContext Lib "imm32" (ByVal hIMC As LongPtr) As Long
Private Declare PtrSafe Function ImmGetContext Lib "imm32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmReleaseContext Lib "imm32" (ByVal hWnd As LongPtr, ByVal hIMC As LongPtr) As Long
Private Declare PtrSafe Function ImmGetOpenStatus Lib "imm32" (ByVal hIMC As LongPtr) As Long
Private Declare PtrSafe Function ImmSetOpenStatus Lib "imm32" (ByVal hIMC As LongPtr, ByVal fOpen As Long) As Long
Private Declare PtrSafe Function ImmAssociateContext Lib "imm32" (ByVal hWnd As LongPtr, ByVal hIMC As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmGetConversionStatus Lib "imm32" (ByVal hIMC As LongPtr, ByVal lpfdwConversion As LongPtr, ByVal lpfdwSentence As LongPtr) As Long
Private Declare PtrSafe Function ImmSetConversionStatus Lib "imm32" (ByVal hIMC As LongPtr, ByVal fdwConversion As Long, ByVal fdwSentence As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare PtrSafe Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare PtrSafe Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare PtrSafe Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare PtrSafe Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long
Private Declare PtrSafe Function IsDialogMessage Lib "user32" Alias "IsDialogMessageW" (ByVal hDlg As LongPtr, ByRef lpMsg As TMSG) As Long
Private Declare PtrSafe Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As Any) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPTr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Declare Function ComCtlsObjAddRef Lib "msvbvm60.dll" Alias "__vbaObjAddref" (ByVal lpObject As Long) As Long
Public Declare Function ComCtlsObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Public Declare Function ComCtlsObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function MCIWndRegisterClass Lib "msvfw32" () As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As TMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwThreadID As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function ImmIsIME Lib "imm32" (ByVal hKL As Long) As Long
Private Declare Function ImmCreateContext Lib "imm32" () As Long
Private Declare Function ImmDestroyContext Lib "imm32" (ByVal hIMC As Long) As Long
Private Declare Function ImmGetContext Lib "imm32" (ByVal hWnd As Long) As Long
Private Declare Function ImmReleaseContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmGetOpenStatus Lib "imm32" (ByVal hIMC As Long) As Long
Private Declare Function ImmSetOpenStatus Lib "imm32" (ByVal hIMC As Long, ByVal fOpen As Long) As Long
Private Declare Function ImmAssociateContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmGetConversionStatus Lib "imm32" (ByVal hIMC As Long, ByVal lpfdwConversion As Long, ByVal lpfdwSentence As Long) As Long
Private Declare Function ImmSetConversionStatus Lib "imm32" (ByVal hIMC As Long, ByVal fdwConversion As Long, ByVal fdwSentence As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageW" (ByVal hDlg As Long, ByRef lpMsg As TMSG) As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclassW2K Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclassW2K Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProcW2K Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const GCW_ATOM As Long = (-32)
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private Const WM_INITDIALOG As Long = &H110
Private Const WM_USER As Long = &H400
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0
Private ShellModHandle As LongPtr, ShellModCount As Long
Private ComCtlsSubclassProcPtr As LongPtr
#If (VBA7 = 0) Then
Private ComCtlsSubclassW2K As Integer
#End If
Private MCIWndRefCount As Long
Private CdlPDEXVTableIPDCB(0 To 5) As LongPtr
Private CdlFRHookHandle As LongPtr
Private CdlFRDialogHandle() As LongPtr, CdlFRDialogCount As Long

#If ImplementPreTranslateMsg = True Then

Private Const UM_PRETRANSLATEMSG As Long = (WM_USER + 1100)
Private ComCtlsPreTranslateMsgHookHandle As LongPtr
Private ComCtlsPreTranslateMsgHwnd As LongPtr, ComCtlsPreTranslateMsgCount As Long

#End If

Public Sub ComCtlsLoadShellMod()
If ShellModHandle = NULL_PTR And ShellModCount = 0 Then ShellModHandle = LoadLibrary(StrPtr("shell32.dll"))
ShellModCount = ShellModCount + 1
End Sub

Public Sub ComCtlsReleaseShellMod()
ShellModCount = ShellModCount - 1
If ShellModHandle <> NULL_PTR And ShellModCount = 0 Then
    FreeLibrary ShellModHandle
    ShellModHandle = NULL_PTR
End If
End Sub

Public Sub ComCtlsInitCC(ByVal ICC As Long)
Dim ICCEX As TINITCOMMONCONTROLSEX
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC
End With
InitCommonControlsEx ICCEX
End Sub

#If VBA7 Then
Public Sub ComCtlsShowAllUIStates(ByVal hWnd As LongPtr)
#Else
Public Sub ComCtlsShowAllUIStates(ByVal hWnd As Long)
#End If
Const WM_UPDATEUISTATE As Long = &H128
Const UIS_CLEAR As Long = 2, UISF_HIDEFOCUS As Long = &H1, UISF_HIDEACCEL As Long = &H2
SendMessage hWnd, WM_UPDATEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS Or UISF_HIDEACCEL), ByVal 0&
End Sub

Public Sub ComCtlsInitBorderStyle(ByRef dwStyle As Long, ByRef dwExStyle As Long, ByVal Value As CCBorderStyleConstants)
Const WS_BORDER As Long = &H800000, WS_DLGFRAME As Long = &H400000
Const WS_EX_CLIENTEDGE As Long = &H200, WS_EX_STATICEDGE As Long = &H20000, WS_EX_WINDOWEDGE As Long = &H100
Select Case Value
    Case CCBorderStyleSingle
        dwStyle = dwStyle Or WS_BORDER
    Case CCBorderStyleThin
        dwExStyle = dwExStyle Or WS_EX_STATICEDGE
    Case CCBorderStyleSunken
        dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
    Case CCBorderStyleRaised
        dwExStyle = dwExStyle Or WS_EX_WINDOWEDGE
        dwStyle = dwStyle Or WS_DLGFRAME
End Select
End Sub

#If VBA7 Then
Public Sub ComCtlsChangeBorderStyle(ByVal hWnd As LongPtr, ByVal Value As CCBorderStyleConstants)
#Else
Public Sub ComCtlsChangeBorderStyle(ByVal hWnd As Long, ByVal Value As CCBorderStyleConstants)
#End If
Const WS_BORDER As Long = &H800000, WS_DLGFRAME As Long = &H400000
Const WS_EX_CLIENTEDGE As Long = &H200, WS_EX_STATICEDGE As Long = &H20000, WS_EX_WINDOWEDGE As Long = &H100
Dim dwStyle As Long, dwExStyle As Long
dwStyle = GetWindowLong(hWnd, GWL_STYLE)
dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
If (dwStyle And WS_BORDER) = WS_BORDER Then dwStyle = dwStyle And Not WS_BORDER
If (dwStyle And WS_DLGFRAME) = WS_DLGFRAME Then dwStyle = dwStyle And Not WS_DLGFRAME
If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
If (dwExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then dwExStyle = dwExStyle And Not WS_EX_WINDOWEDGE
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, Value)
SetWindowLong hWnd, GWL_STYLE, dwStyle
SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
Call ComCtlsFrameChanged(hWnd)
End Sub

#If VBA7 Then
Public Sub ComCtlsFrameChanged(ByVal hWnd As LongPtr)
#Else
Public Sub ComCtlsFrameChanged(ByVal hWnd As Long)
#End If
Const SWP_FRAMECHANGED As Long = &H20, SWP_NOMOVE As Long = &H2, SWP_NOOWNERZORDER As Long = &H200, SWP_NOSIZE As Long = &H1, SWP_NOZORDER As Long = &H4
SetWindowPos hWnd, NULL_PTR, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

#If VBA7 Then
Public Sub ComCtlsInitToolTip(ByVal hWnd As LongPtr)
#Else
Public Sub ComCtlsInitToolTip(ByVal hWnd As Long)
#End If
#If VBA7 Then
Const HWND_TOPMOST As LongPtr = (-1)
#Else
Const HWND_TOPMOST As Long = (-1)
#End If
Const WS_EX_TOPMOST As Long = &H8
Const SWP_NOMOVE As Long = &H2, SWP_NOSIZE As Long = &H1, SWP_NOACTIVATE As Long = &H10
If Not (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOPMOST) = WS_EX_TOPMOST Then SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
SendMessage hWnd, TTM_SETMAXTIPWIDTH, 0, ByVal &H7FFF&
End Sub

#If VBA7 Then
Public Sub ComCtlsCreateIMC(ByVal hWnd As LongPtr, ByRef hIMC As LongPtr)
#Else
Public Sub ComCtlsCreateIMC(ByVal hWnd As Long, ByRef hIMC As Long)
#End If
If hIMC = NULL_PTR Then
    hIMC = ImmCreateContext()
    If hIMC <> NULL_PTR Then ImmAssociateContext hWnd, hIMC
End If
End Sub

#If VBA7 Then
Public Sub ComCtlsDestroyIMC(ByVal hWnd As LongPtr, ByRef hIMC As LongPtr)
#Else
Public Sub ComCtlsDestroyIMC(ByVal hWnd As Long, ByRef hIMC As Long)
#End If
If hIMC <> NULL_PTR Then
    ImmAssociateContext hWnd, NULL_PTR
    ImmDestroyContext hIMC
    hIMC = NULL_PTR
End If
End Sub

#If VBA7 Then
Public Sub ComCtlsSetIMEMode(ByVal hWnd As LongPtr, ByVal hIMCOrig As LongPtr, ByVal Value As CCIMEModeConstants)
#Else
Public Sub ComCtlsSetIMEMode(ByVal hWnd As Long, ByVal hIMCOrig As Long, ByVal Value As CCIMEModeConstants)
#End If
Const IME_CMODE_ALPHANUMERIC As Long = &H0, IME_CMODE_NATIVE As Long = &H1, IME_CMODE_KATAKANA As Long = &H2, IME_CMODE_FULLSHAPE As Long = &H8
Dim hKL As LongPtr
hKL = GetKeyboardLayout(0)
If ImmIsIME(hKL) = NULL_PTR Or hIMCOrig = NULL_PTR Then Exit Sub
Dim hIMC As LongPtr
hIMC = ImmGetContext(hWnd)
If Value = CCIMEModeDisable Then
    If hIMC <> NULL_PTR Then
        ImmReleaseContext hWnd, hIMC
        ImmAssociateContext hWnd, NULL_PTR
    End If
Else
    If hIMC = NULL_PTR Then
        ImmAssociateContext hWnd, hIMCOrig
        hIMC = ImmGetContext(hWnd)
    End If
    If hIMC <> NULL_PTR And Value <> CCIMEModeNoControl Then
        Dim dwConversion As Long, dwSentence As Long
        ImmGetConversionStatus hIMC, VarPtr(dwConversion), VarPtr(dwSentence)
        Select Case Value
            Case CCIMEModeOn
                ImmSetOpenStatus hIMC, 1
            Case CCIMEModeOff
                ImmSetOpenStatus hIMC, 0
            Case CCIMEModeHiragana
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
            Case CCIMEModeKatakana
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion Or IME_CMODE_KATAKANA
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
            Case CCIMEModeKatakanaHalf
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion Or IME_CMODE_KATAKANA
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
            Case CCIMEModeAlphaFull
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
                If (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion And Not IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
            Case CCIMEModeAlpha
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_ALPHANUMERIC) = IME_CMODE_ALPHANUMERIC Then dwConversion = dwConversion Or IME_CMODE_ALPHANUMERIC
                If (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion And Not IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
            Case CCIMEModeHangulFull
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
            Case CCIMEModeHangul
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
        End Select
        ImmSetConversionStatus hIMC, dwConversion, dwSentence
        ImmReleaseContext hWnd, hIMC
    End If
End If
End Sub

#If VBA7 Then
Public Sub ComCtlsRequestMouseLeave(ByVal hWnd As LongPtr)
#Else
Public Sub ComCtlsRequestMouseLeave(ByVal hWnd As Long)
#End If
Const TME_LEAVE As Long = &H2
Dim TME As TRACKMOUSEEVENTSTRUCT
With TME
.cbSize = LenB(TME)
.hWndTrack = hWnd
.dwFlags = TME_LEAVE
End With
TrackMouseEvent TME
End Sub

Public Sub ComCtlsCheckRightToLeft(ByRef Value As Boolean, ByVal UserControlValue As Boolean, ByVal ModeValue As CCRightToLeftModeConstants)
If Value = False Then Exit Sub
Select Case ModeValue
    Case CCRightToLeftModeNoControl
    Case CCRightToLeftModeVBAME
        Value = UserControlValue
    Case CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
        Const LOCALE_FONTSIGNATURE As Long = &H58, SORT_DEFAULT As Long = &H0
        Dim LangID As Integer, LCID As Long, LocaleSig As TLOCALESIGNATURE
        Select Case ModeValue
            Case CCRightToLeftModeSystemLocale
                LangID = GetSystemDefaultLangID()
            Case CCRightToLeftModeUserLocale
                LangID = GetUserDefaultLangID()
            Case CCRightToLeftModeOSLanguage
                LangID = GetUserDefaultUILanguage()
        End Select
        LCID = (SORT_DEFAULT * &H10000) Or LangID
        If GetLocaleInfo(LCID, LOCALE_FONTSIGNATURE, VarPtr(LocaleSig), (LenB(LocaleSig) / 2)) <> 0 Then
            ' Unicode subset bitfield 0 to 127. Bit 123 = Layout progress, horizontal from right to left
            Value = CBool((LocaleSig.lsUsb(15) And (2 ^ (4 - 1))) <> 0)
        End If
End Select
End Sub

#If VBA7 Then
Public Sub ComCtlsSetRightToLeft(ByVal hWnd As LongPtr, ByVal dwMask As Long)
#Else
Public Sub ComCtlsSetRightToLeft(ByVal hWnd As Long, ByVal dwMask As Long)
#End If
Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
' WS_EX_LAYOUTRTL will take care of both layout and reading order with the single flag and mirrors the window.
Dim dwExStyle As Long
dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
If (dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle And Not WS_EX_LAYOUTRTL
If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If (dwMask And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
If (dwMask And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle Or WS_EX_RIGHT
If (dwMask And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle Or WS_EX_LEFTSCROLLBAR
Const WS_POPUP As Long = &H80000000
If (GetWindowLong(hWnd, GWL_STYLE) And WS_POPUP) = 0 Then
    SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
    InvalidateRect hWnd, ByVal NULL_PTR, 1
    Call ComCtlsFrameChanged(hWnd)
Else
    ' ToolTip control supports only the WS_EX_LAYOUTRTL flag.
    ' Set TTF_RTLREADING flag when dwMask contains WS_EX_RTLREADING, though WS_EX_RTLREADING will not be actually set.
    If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
    If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
    If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
    SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
    Const TTM_SETTOOLINFOA As Long = (WM_USER + 9)
    Const TTM_SETTOOLINFOW As Long = (WM_USER + 54)
    Const TTM_SETTOOLINFO As Long = TTM_SETTOOLINFOW
    Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
    Const TTM_ENUMTOOLSA As Long = (WM_USER + 14)
    Const TTM_ENUMTOOLSW As Long = (WM_USER + 58)
    Const TTM_ENUMTOOLS As Long = TTM_ENUMTOOLSW
    Const TTM_UPDATE As Long = (WM_USER + 29)
    Const TTF_RTLREADING As Long = &H4
    Dim i As Long, TI As TOOLINFO, Buffer As String
    With TI
    .cbSize = LenB(TI)
    Buffer = String(80, vbNullChar)
    .lpszText = StrPtr(Buffer)
    For i = 1 To CLng(SendMessage(hWnd, TTM_GETTOOLCOUNT, 0, ByVal 0&))
        If SendMessage(hWnd, TTM_ENUMTOOLS, i - 1, ByVal VarPtr(TI)) <> 0 Then
            If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Or (dwMask And WS_EX_RTLREADING) = 0 Then
                If (.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then .uFlags = .uFlags And Not TTF_RTLREADING
            Else
                If (.uFlags And TTF_RTLREADING) = 0 Then .uFlags = .uFlags Or TTF_RTLREADING
            End If
            SendMessage hWnd, TTM_SETTOOLINFO, 0, ByVal VarPtr(TI)
            SendMessage hWnd, TTM_UPDATE, 0, ByVal 0&
        End If
    Next i
    End With
End If
End Sub

Public Sub ComCtlsIPPBSetPredefinedStringsImageList(ByRef StringsOut() As String, ByRef CookiesOut() As Long, ByRef ControlsEnum As VBRUN.ParentControls, ByRef ImageListArray() As String)
Dim ControlEnum As Object, PropUBound As Long
PropUBound = UBound(StringsOut())
ReDim Preserve StringsOut(PropUBound + 1) As String
ReDim Preserve CookiesOut(PropUBound + 1) As Long
StringsOut(PropUBound) = "(None)"
CookiesOut(PropUBound) = PropUBound
For Each ControlEnum In ControlsEnum
    If TypeName(ControlEnum) = "ImageList" Then
        PropUBound = UBound(StringsOut())
        ReDim Preserve StringsOut(PropUBound + 1) As String
        ReDim Preserve CookiesOut(PropUBound + 1) As Long
        StringsOut(PropUBound) = ProperControlName(ControlEnum)
        CookiesOut(PropUBound) = PropUBound
    End If
Next ControlEnum
PropUBound = UBound(StringsOut())
ReDim ImageListArray(0 To PropUBound) As String
Dim i As Long
For i = 0 To PropUBound
    ImageListArray(i) = StringsOut(i)
Next i
End Sub

Public Sub ComCtlsPPInitComboMousePointer(ByVal ComboBox As Object)
With ComboBox
.AddItem CCMousePointerDefault & " - Default"
.ItemData(.NewIndex) = CCMousePointerDefault
.AddItem CCMousePointerArrow & " - Arrow"
.ItemData(.NewIndex) = CCMousePointerArrow
.AddItem CCMousePointerCrosshair & " - Cross"
.ItemData(.NewIndex) = CCMousePointerCrosshair
.AddItem CCMousePointerIbeam & " - I-Beam"
.ItemData(.NewIndex) = CCMousePointerIbeam
.AddItem CCMousePointerHand & " - Hand"
.ItemData(.NewIndex) = CCMousePointerHand
.AddItem CCMousePointerSizePointer & " - Size"
.ItemData(.NewIndex) = CCMousePointerSizePointer
.AddItem CCMousePointerSizeNESW & " - Size NE SW"
.ItemData(.NewIndex) = CCMousePointerSizeNESW
.AddItem CCMousePointerSizeNS & " - Size N S"
.ItemData(.NewIndex) = CCMousePointerSizeNS
.AddItem CCMousePointerSizeNWSE & " - Size NW SE"
.ItemData(.NewIndex) = CCMousePointerSizeNWSE
.AddItem CCMousePointerSizeWE & " - Size W E"
.ItemData(.NewIndex) = CCMousePointerSizeWE
.AddItem CCMousePointerUpArrow & " - Up Arrow"
.ItemData(.NewIndex) = CCMousePointerUpArrow
.AddItem CCMousePointerHourglass & " - Hourglass"
.ItemData(.NewIndex) = CCMousePointerHourglass
.AddItem CCMousePointerNoDrop & " - No Drop"
.ItemData(.NewIndex) = CCMousePointerNoDrop
.AddItem CCMousePointerArrowHourglass & " - Arrow and Hourglass"
.ItemData(.NewIndex) = CCMousePointerArrowHourglass
.AddItem CCMousePointerArrowQuestion & " - Arrow and Question"
.ItemData(.NewIndex) = CCMousePointerArrowQuestion
.AddItem CCMousePointerSizeAll & " - Size All"
.ItemData(.NewIndex) = CCMousePointerSizeAll
.AddItem CCMousePointerArrowCD & " - Arrow and CD"
.ItemData(.NewIndex) = CCMousePointerArrowCD
.AddItem CCMousePointerCustom & " - Custom"
.ItemData(.NewIndex) = CCMousePointerCustom
End With
End Sub

Public Sub ComCtlsPPInitComboIMEMode(ByVal ComboBox As Object)
With ComboBox
.AddItem CCIMEModeNoControl & " - NoControl"
.ItemData(.NewIndex) = CCIMEModeNoControl
.AddItem CCIMEModeOn & " - On"
.ItemData(.NewIndex) = CCIMEModeOn
.AddItem CCIMEModeOff & " - Off"
.ItemData(.NewIndex) = CCIMEModeOff
.AddItem CCIMEModeDisable & " - Disable"
.ItemData(.NewIndex) = CCIMEModeDisable
.AddItem CCIMEModeHiragana & " - Hiragana"
.ItemData(.NewIndex) = CCIMEModeHiragana
.AddItem CCIMEModeKatakana & " - Katakana"
.ItemData(.NewIndex) = CCIMEModeKatakana
.AddItem CCIMEModeKatakanaHalf & " - KatakanaHalf"
.ItemData(.NewIndex) = CCIMEModeKatakanaHalf
.AddItem CCIMEModeAlphaFull & " - AlphaFull"
.ItemData(.NewIndex) = CCIMEModeAlphaFull
.AddItem CCIMEModeAlpha & " - Alpha"
.ItemData(.NewIndex) = CCIMEModeAlpha
.AddItem CCIMEModeHangulFull & " - HangulFull"
.ItemData(.NewIndex) = CCIMEModeHangulFull
.AddItem CCIMEModeHangul & " - Hangul"
.ItemData(.NewIndex) = CCIMEModeHangul
End With
End Sub

Public Sub ComCtlsPPKeyPressOnlyNumeric(ByRef KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then If KeyAscii <> 8 Then KeyAscii = 0
End Sub

#If VBA7 Then
Public Function ComCtlsPeekCharCode(ByVal hWnd As LongPtr) As Long
#Else
Public Function ComCtlsPeekCharCode(ByVal hWnd As Long) As Long
#End If
Dim Msg As TMSG
Const PM_NOREMOVE As Long = &H0, WM_CHAR As Long = &H102
If PeekMessage(Msg, hWnd, WM_CHAR, WM_CHAR, PM_NOREMOVE) <> 0 Then ComCtlsPeekCharCode = CLng(Msg.wParam)
End Function

Public Function ComCtlsSupportLevel() As Integer
Static Done As Boolean, Value As Integer
If Done = False Then
    Dim Version As DLLVERSIONINFO
    On Error Resume Next
    Version.cbSize = LenB(Version)
    If DllGetVersion(Version) = S_OK Then
        If Version.dwMajor = 6 And Version.dwMinor = 0 Then
            Value = 1
        ElseIf Version.dwMajor > 6 Or (Version.dwMajor = 6 And Version.dwMinor > 0) Then
            Value = 2
        End If
    End If
    Done = True
End If
ComCtlsSupportLevel = Value
End Function

#If VBA7 Then
Public Sub ComCtlsSetSubclass(ByVal hWnd As LongPtr, ByVal This As ISubclass, ByVal dwRefData As LongPtr, Optional ByVal Name As String)
#Else
Public Sub ComCtlsSetSubclass(ByVal hWnd As Long, ByVal This As ISubclass, ByVal dwRefData As Long, Optional ByVal Name As String)
#End If
If hWnd = NULL_PTR Then Exit Sub
If Name = vbNullString Then Name = "ComCtls"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If ComCtlsSubclassProcPtr = NULL_PTR Then ComCtlsSubclassProcPtr = ProcPtr(AddressOf ComCtlsSubclassProc)
    #If VBA7 Then
    SetWindowSubclass hWnd, ComCtlsSubclassProcPtr, ObjPtr(This), dwRefData
    #Else
    If ComCtlsSubclassW2K = 0 Then
        Dim hLib As Long
        hLib = LoadLibrary(StrPtr("comctl32.dll"))
        If hLib <> NULL_PTR Then
            If GetProcAddress(hLib, "SetWindowSubclass") <> NULL_PTR Then
                ComCtlsSubclassW2K = 1
            ElseIf GetProcAddress(hLib, 410&) <> NULL_PTR Then
                ComCtlsSubclassW2K = -1
            End If
            FreeLibrary hLib
        End If
    End If
    If ComCtlsSubclassW2K > -1 Then
        SetWindowSubclass hWnd, ComCtlsSubclassProcPtr, ObjPtr(This), dwRefData
    Else
        SetWindowSubclassW2K hWnd, ComCtlsSubclassProcPtr, ObjPtr(This), dwRefData
    End If
    #End If
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

#If VBA7 Then
Public Function ComCtlsDefaultProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsDefaultProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
#If VBA7 Then
ComCtlsDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
#Else
If ComCtlsSubclassW2K > -1 Then
    ComCtlsDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
Else
    ComCtlsDefaultProc = DefSubclassProcW2K(hWnd, wMsg, wParam, lParam)
End If
#End If
End Function

#If VBA7 Then
Public Sub ComCtlsRemoveSubclass(ByVal hWnd As LongPtr, Optional ByVal Name As String)
#Else
Public Sub ComCtlsRemoveSubclass(ByVal hWnd As Long, Optional ByVal Name As String)
#End If
If hWnd = NULL_PTR Then Exit Sub
If Name = vbNullString Then Name = "ComCtls"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    #If VBA7 Then
    RemoveWindowSubclass hWnd, ComCtlsSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    #Else
    If ComCtlsSubclassW2K > -1 Then
        RemoveWindowSubclass hWnd, ComCtlsSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    Else
        RemoveWindowSubclassW2K hWnd, ComCtlsSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    End If
    #End If
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

#If VBA7 Then
Public Function ComCtlsSubclassProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Public Function ComCtlsSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
#End If
Select Case wMsg
    Case WM_DESTROY
        ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        #If VBA7 Then
        RemoveWindowSubclass hWnd, ComCtlsSubclassProcPtr, uIdSubclass
        #Else
        If ComCtlsSubclassW2K > -1 Then
            RemoveWindowSubclass hWnd, ComCtlsSubclassProcPtr, uIdSubclass
        Else
            RemoveWindowSubclassW2K hWnd, ComCtlsSubclassProcPtr, uIdSubclass
        End If
        #End If
        Exit Function
End Select
On Error Resume Next
Dim This As ISubclass
Set This = PtrToObj(uIdSubclass)
If Err.Number = 0 Then
    ComCtlsSubclassProc = This.Message(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub ComCtlsImlListImageIndex(ByVal Control As Object, ByVal ImageList As Variant, ByVal KeyOrIndex As Variant, ByRef ImageIndex As Long)
Dim LngValue As Long
Select Case VarType(KeyOrIndex)
    Case vbLong, vbInteger, vbByte
        LngValue = KeyOrIndex
    Case vbString
        Dim ImageListControl As Object
        If IsObject(ImageList) Then
            Set ImageListControl = ImageList
        ElseIf VarType(ImageList) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In Control.ControlsEnum
                If TypeName(ControlEnum) = "ImageList" Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = ImageList And Not CompareName = vbNullString Then
                        Set ImageListControl = ControlEnum
                        Exit For
                    End If
                End If
            Next ControlEnum
        End If
        If Not ImageListControl Is Nothing Then
            On Error Resume Next
            LngValue = ImageListControl.ListImages(KeyOrIndex).Index
            On Error GoTo 0
        End If
        If LngValue = 0 Then Err.Raise Number:=35601, Description:="Element not found"
    Case vbDouble, vbSingle
        LngValue = CLng(KeyOrIndex)
    Case vbEmpty
    Case Else
        Err.Raise 13
End Select
If LngValue < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
ImageIndex = LngValue
End Sub

Public Sub ComCtlsMCIWndRegisterClass()
If MCIWndRefCount = 0 Then MCIWndRegisterClass
MCIWndRefCount = MCIWndRefCount + 1
End Sub

Public Sub ComCtlsMCIWndReleaseClass()
MCIWndRefCount = MCIWndRefCount - 1
If MCIWndRefCount = 0 Then UnregisterClass StrPtr("MCIWndClass"), App.hInstance
End Sub

#If VBA7 Then
Public Function ComCtlsCbrPlaceholderWindowProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCbrPlaceholderWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
ComCtlsCbrPlaceholderWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End Function

#If VBA7 Then
Public Function ComCtlsTbrEnumThreadWndProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function ComCtlsTbrEnumThreadWndProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
#End If
If GetClassLong(hWnd, GCW_ATOM) = &H8000& Then ComCtlsTbrEnumThreadWndProc = 0 Else ComCtlsTbrEnumThreadWndProc = 1
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionBinary(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionBinary = CLng(This.Message(0, 0, lParam1, lParam2, 10))
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionText(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionText = CLng(This.Message(0, 0, lParam1, lParam2, 11))
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionNumeric(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionNumeric(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionNumeric = CLng(This.Message(0, 0, lParam1, lParam2, 12))
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionCurrency(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionCurrency(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionCurrency = CLng(This.Message(0, 0, lParam1, lParam2, 13))
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionDate(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionDate(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionDate = CLng(This.Message(0, 0, lParam1, lParam2, 14))
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionLogical(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionLogical(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionLogical = CLng(This.Message(0, 0, lParam1, lParam2, 15))
End Function

#If VBA7 Then
Public Function ComCtlsLvwSortingFunctionGroups(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsLvwSortingFunctionGroups(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsLvwSortingFunctionGroups = CLng(This.Message(0, 0, lParam1, lParam2, 0))
End Function

#If VBA7 Then
Public Function ComCtlsTvwSortingFunctionBinary(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsTvwSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsTvwSortingFunctionBinary = CLng(This.Message(0, 0, lParam1, lParam2, 10))
End Function

#If VBA7 Then
Public Function ComCtlsTvwSortingFunctionText(ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsTvwSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsTvwSortingFunctionText = CLng(This.Message(0, 0, lParam1, lParam2, 11))
End Function

#If VBA7 Then
Public Function ComCtlsFtcEnumFontFunction(ByVal lpELF As LongPtr, ByVal lpTM As LongPtr, ByVal FontType As Long, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsFtcEnumFontFunction(ByVal lpELF As Long, ByVal lpTM As Long, ByVal FontType As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsFtcEnumFontFunction = CLng(This.Message(0, FontType, lpELF, lpTM, 10))
End Function

#If VBA7 Then
Public Function ComCtlsCdlOFN1CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlOFN1CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN1CallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -1)
Else
    ComCtlsCdlOFN1CallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlOFN1CallbackProcOldStyle(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlOFN1CallbackProcOldStyle(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcOldStyleCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("ComCtlsCdlOFN1CallbackProcOldStyleCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN1CallbackProcOldStyle = This.Message(hDlg, wMsg, wParam, lParam, -1001)
Else
    ComCtlsCdlOFN1CallbackProcOldStyle = 0
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlOFN2CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlOFN2CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN2CallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -2)
Else
    ComCtlsCdlOFN2CallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlOFN2CallbackProcOldStyle(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlOFN2CallbackProcOldStyle(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcOldStyleCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("ComCtlsCdlOFN2CallbackProcOldStyleCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlOFN2CallbackProcOldStyle = This.Message(hDlg, wMsg, wParam, lParam, -1002)
Else
    ComCtlsCdlOFN2CallbackProcOldStyle = 0
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlCCCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlCCCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlCCCallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 48), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 24), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("ComCtlsCdlCCCallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlCCCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -3)
Else
    ComCtlsCdlCCCallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlCFCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlCFCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("ComCtlsCdlCFCallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 40), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("ComCtlsCdlCFCallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As ISubclass
    Set This = PtrToObj(lCustData)
    ComCtlsCdlCFCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -4)
Else
    ComCtlsCdlCFCallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlPDCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlPDCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlPDCallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 54), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 38), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        ComCtlsCdlPDCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -5)
    Else
        ComCtlsCdlPDCallbackProc = 0
    End If
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlPDEXCallbackPtr(ByVal This As ISubclass) As LongPtr
#Else
Public Function ComCtlsCdlPDEXCallbackPtr(ByVal This As ISubclass) As Long
#End If
Dim VTableData(0 To 2) As LongPtr
VTableData(0) = GetVTableIPDCB()
VTableData(1) = 0 ' RefCount is uninstantiated
VTableData(2) = ObjPtr(This)
Dim hMem As LongPtr
hMem = CoTaskMemAlloc(12)
If hMem <> NULL_PTR Then
    CopyMemory ByVal hMem, VTableData(0), 3 * PTR_SIZE
    ComCtlsCdlPDEXCallbackPtr = hMem
End If
End Function

Private Function GetVTableIPDCB() As LongPtr
If CdlPDEXVTableIPDCB(0) = NULL_PTR Then
    CdlPDEXVTableIPDCB(0) = ProcPtr(AddressOf IPDCB_QueryInterface)
    CdlPDEXVTableIPDCB(1) = ProcPtr(AddressOf IPDCB_AddRef)
    CdlPDEXVTableIPDCB(2) = ProcPtr(AddressOf IPDCB_Release)
    CdlPDEXVTableIPDCB(3) = ProcPtr(AddressOf IPDCB_InitDone)
    CdlPDEXVTableIPDCB(4) = ProcPtr(AddressOf IPDCB_SelectionChange)
    CdlPDEXVTableIPDCB(5) = ProcPtr(AddressOf IPDCB_HandleMessage)
End If
GetVTableIPDCB = VarPtr(CdlPDEXVTableIPDCB(0))
End Function

Private Function IPDCB_QueryInterface(ByVal Ptr As LongPtr, ByRef IID As CLSID, ByRef pvObj As LongPtr) As Long
If VarPtr(pvObj) = NULL_PTR Then
    IPDCB_QueryInterface = E_POINTER
    Exit Function
End If
' IID_IPrintDialogCallback = {5852A2C3-6530-11D1-B6A3-0000F8757BF9}
If IID.Data1 = &H5852A2C3 And IID.Data2 = &H6530 And IID.Data3 = &H11D1 Then
    If IID.Data4(0) = &HB6 And IID.Data4(1) = &HA3 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
    And IID.Data4(4) = &HF8 And IID.Data4(5) = &H75 And IID.Data4(6) = &H7B And IID.Data4(7) = &HF9 Then
        pvObj = Ptr
        IPDCB_AddRef Ptr
        IPDCB_QueryInterface = S_OK
    Else
        IPDCB_QueryInterface = E_NOINTERFACE
    End If
Else
    IPDCB_QueryInterface = E_NOINTERFACE
End If
End Function

Private Function IPDCB_AddRef(ByVal Ptr As LongPtr) As Long
CopyMemory IPDCB_AddRef, ByVal UnsignedAdd(Ptr, 1 * PTR_SIZE), 4
IPDCB_AddRef = IPDCB_AddRef + 1
CopyMemory ByVal UnsignedAdd(Ptr, 1 * PTR_SIZE), IPDCB_AddRef, 4
End Function

Private Function IPDCB_Release(ByVal Ptr As LongPtr) As Long
CopyMemory IPDCB_Release, ByVal UnsignedAdd(Ptr, 1 * PTR_SIZE), 4
IPDCB_Release = IPDCB_Release - 1
CopyMemory ByVal UnsignedAdd(Ptr, 1 * PTR_SIZE), IPDCB_Release, 4
If IPDCB_Release = 0 Then CoTaskMemFree Ptr
End Function

Private Function IPDCB_InitDone(ByVal Ptr As LongPtr) As Long
IPDCB_InitDone = S_FALSE
End Function

Private Function IPDCB_SelectionChange(ByVal Ptr As LongPtr) As Long
IPDCB_SelectionChange = S_FALSE
End Function

Private Function IPDCB_HandleMessage(ByVal Ptr As LongPtr, ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByRef Result As LongPtr) As Long
If wMsg = WM_INITDIALOG Then
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(Ptr, 16), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(Ptr, 8), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        This.Message hDlg, wMsg, wParam, lParam, -5
    End If
End If
IPDCB_HandleMessage = S_FALSE
End Function

#If VBA7 Then
Public Function ComCtlsCdlPSDCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlPSDCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlPSDCallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 88), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        ComCtlsCdlPSDCallbackProc = This.Message(hDlg, wMsg, wParam, lParam, -7)
    Else
        ComCtlsCdlPSDCallbackProc = 0
    End If
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlBIFCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal lParam As LongPtr, ByVal This As ISubclass) As Long
#Else
Public Function ComCtlsCdlBIFCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal lParam As Long, ByVal This As ISubclass) As Long
#End If
ComCtlsCdlBIFCallbackProc = CLng(This.Message(hDlg, wMsg, 0, lParam, -8))
End Function

#If VBA7 Then
Public Function ComCtlsCdlFR1CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlFR1CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlFR1CallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 56), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        This.Message hDlg, wMsg, wParam, lParam, -9
    End If
    ' Need to return a nonzero value or else the dialog box will not be shown.
    ComCtlsCdlFR1CallbackProc = 1
End If
End Function

#If VBA7 Then
Public Function ComCtlsCdlFR2CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function ComCtlsCdlFR2CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    ComCtlsCdlFR2CallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 56), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As ISubclass
        Set This = PtrToObj(lCustData)
        This.Message hDlg, wMsg, wParam, lParam, -10
    End If
    ' Need to return a nonzero value or else the dialog box will not be shown.
    ComCtlsCdlFR2CallbackProc = 1
End If
End Function

#If VBA7 Then
Public Sub ComCtlsCdlFRAddHook(ByVal hDlg As LongPtr)
#Else
Public Sub ComCtlsCdlFRAddHook(ByVal hDlg As Long)
#End If
If CdlFRHookHandle = NULL_PTR And CdlFRDialogCount = 0 Then
    Const WH_GETMESSAGE As Long = 3
    CdlFRHookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf ComCtlsCdlFRHookProc, NULL_PTR, App.ThreadID)
    ReDim CdlFRDialogHandle(0) ' As LongPtr
    CdlFRDialogHandle(0) = hDlg
Else
    ReDim Preserve CdlFRDialogHandle(0 To CdlFRDialogCount) ' As LongPtr
    CdlFRDialogHandle(CdlFRDialogCount) = hDlg
End If
CdlFRDialogCount = CdlFRDialogCount + 1
End Sub

#If VBA7 Then
Public Sub ComCtlsCdlFRReleaseHook(ByVal hDlg As LongPtr)
#Else
Public Sub ComCtlsCdlFRReleaseHook(ByVal hDlg As Long)
#End If
Dim Index As Long, i As Long
Index = -1
For i = 0 To CdlFRDialogCount - 1
    If CdlFRDialogHandle(i) = hDlg Then
        Index = i
        Exit For
    End If
Next i
If Index > -1 Then
    CdlFRDialogCount = CdlFRDialogCount - 1
    If CdlFRHookHandle <> NULL_PTR And CdlFRDialogCount = 0 Then
        UnhookWindowsHookEx CdlFRHookHandle
        CdlFRHookHandle = NULL_PTR
        Erase CdlFRDialogHandle()
    Else
        If Index < CdlFRDialogCount Then
            For i = Index To CdlFRDialogCount - 1
                CdlFRDialogHandle(i) = CdlFRDialogHandle(i + 1)
            Next i
        End If
        ReDim Preserve CdlFRDialogHandle(0 To CdlFRDialogCount - 1) ' As LongPtr
    End If
End If
End Sub

Private Function ComCtlsCdlFRHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Const HC_ACTION As Long = 0, PM_REMOVE As Long = &H1
Const WM_KEYFIRST As Long = &H100, WM_KEYLAST As Long = &H108, WM_NULL As Long = &H0
If nCode >= HC_ACTION And wParam = PM_REMOVE Then
    Dim Msg As TMSG
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    If Msg.Message >= WM_KEYFIRST And Msg.Message <= WM_KEYLAST Then
        If CdlFRDialogCount > 0 Then
            Dim i As Long
            For i = 0 To CdlFRDialogCount - 1
                If IsDialogMessage(CdlFRDialogHandle(i), Msg) <> 0 Then
                    Msg.Message = WM_NULL
                    Msg.wParam = 0
                    Msg.lParam = 0
                    CopyMemory ByVal lParam, Msg, LenB(Msg)
                    Exit For
                End If
            Next i
        End If
    End If
End If
ComCtlsCdlFRHookProc = CallNextHookEx(CdlFRHookHandle, nCode, wParam, lParam)
End Function

#If ImplementPreTranslateMsg = True Then

Public Sub ComCtlsPreTranslateMsgAddHook()
If ComCtlsPreTranslateMsgHookHandle = NULL_PTR And ComCtlsPreTranslateMsgCount = 0 Then
    Const WH_GETMESSAGE As Long = 3
    ComCtlsPreTranslateMsgHookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf ComCtlsPreTranslateMsgHookProc, NULL_PTR, App.ThreadID)
End If
ComCtlsPreTranslateMsgCount = ComCtlsPreTranslateMsgCount + 1
End Sub

Public Sub ComCtlsPreTranslateMsgReleaseHook()
ComCtlsPreTranslateMsgCount = ComCtlsPreTranslateMsgCount - 1
If ComCtlsPreTranslateMsgHookHandle <> NULL_PTR And ComCtlsPreTranslateMsgCount = 0 Then
    UnhookWindowsHookEx ComCtlsPreTranslateMsgHookHandle
    ComCtlsPreTranslateMsgHookHandle = NULL_PTR
    ComCtlsPreTranslateMsgHwnd = NULL_PTR
End If
End Sub

#If VBA7 Then
Public Sub ComCtlsPreTranslateMsgActivate(ByVal hWnd As LongPtr)
#Else
Public Sub ComCtlsPreTranslateMsgActivate(ByVal hWnd As Long)
#End If
ComCtlsPreTranslateMsgHwnd = hWnd
End Sub

Public Sub ComCtlsPreTranslateMsgDeActivate()
ComCtlsPreTranslateMsgHwnd = NULL_PTR
End Sub

Private Function ComCtlsPreTranslateMsgHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Const HC_ACTION As Long = 0, PM_REMOVE As Long = &H1
Const WM_KEYFIRST As Long = &H100, WM_KEYLAST As Long = &H108, WM_NULL As Long = &H0
If nCode >= HC_ACTION And wParam = PM_REMOVE Then
    Dim Msg As TMSG
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    If Msg.Message >= WM_KEYFIRST And Msg.Message <= WM_KEYLAST Then
        If ComCtlsPreTranslateMsgHwnd <> NULL_PTR And ComCtlsPreTranslateMsgCount > 0 Then
            If Msg.hWnd = ComCtlsPreTranslateMsgHwnd Then
                If SendMessage(Msg.hWnd, UM_PRETRANSLATEMSG, 0, ByVal lParam) <> 0 Then
                    Msg.Message = WM_NULL
                    Msg.wParam = 0
                    Msg.lParam = 0
                    CopyMemory ByVal lParam, Msg, LenB(Msg)
                End If
            End If
        End If
    End If
End If
ComCtlsPreTranslateMsgHookProc = CallNextHookEx(ComCtlsPreTranslateMsgHookHandle, nCode, wParam, lParam)
End Function

#End If
