VERSION 5.00
Begin VB.UserControl LinkLabel 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "LinkLabel.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "LinkLabel.ctx":004A
End
Attribute VB_Name = "LinkLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

#If False Then
Private LlbLinkActivateReasonClick, LlbLinkActivateReasonReturn
#End If
Public Enum LlbLinkActivateReasonConstants
LlbLinkActivateReasonClick = 0
LlbLinkActivateReasonReturn = 1
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type TMSG
hWnd As LongPtr
Message As Long
wParam As LongPtr
lParam As LongPtr
Time As Long
PT As POINTAPI
End Type
Private Const L_MAX_URL_LENGTH As Long = 2084
Private Const MAX_LINKID_TEXT As Long = 48
Private Type LITEM
Mask As Long
iLink As Long
State As Long
StateMask As Long
szID(0 To ((MAX_LINKID_TEXT * 2) - 1)) As Byte
szURL(0 To ((L_MAX_URL_LENGTH * 2) - 1)) As Byte
End Type
Private Type LHITTESTINFO
PT As POINTAPI
Item As LITEM
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
Private Type NMHDR
hWndFrom As LongPtr
IDFrom As LongPtr
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_NEWFONT As Long = &H2
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As LongPtr
RC As RECT
dwItemSpec As LongPtr
uItemState As Long
lItemlParam As LongPtr
End Type
Private Type NMLINK
hdr As NMHDR
Item As LITEM
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As LongPtr
szText(0 To ((80 * 2) - 1)) As Byte
hInst As LongPtr
uFlags As Long
lParam As LongPtr
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event LinkActivate(ByVal Link As LlbLink, ByVal Reason As LlbLinkActivateReasonConstants)
Attribute LinkActivate.VB_Description = "Occurs when a link item is activated."
Public Event LinkForeColor(ByVal Link As LlbLink, ByRef RGBColor As Long)
Attribute LinkForeColor.VB_Description = "Occurs when a link item is about to draw the text. This is a request to provide an alternative foreground color. The foreground color is passed in an RGB format. Requires comctl32.dll version 6.1 or higher."
Public Event LinkGetTipText(ByVal Link As LlbLink, ByRef Text As String)
Attribute LinkGetTipText.VB_Description = "Occurs if the link label control is about to display a tool tip on a link item and requests the text to display. This will only occur if the show tips property is set to true."
Public Event LinkMouseEnter(ByVal Link As LlbLink)
Attribute LinkMouseEnter.VB_Description = "Occurs when the user moves the mouse into a link item."
Public Event LinkMouseLeave(ByVal Link As LlbLink)
Attribute LinkMouseLeave.VB_Description = "Occurs when the user moves the mouse out of a link item."
Public Event PreviewKeyDown(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyDown.VB_Description = "Occurs before the KeyDown event."
Public Event PreviewKeyUp(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyUp.VB_Description = "Occurs before the KeyUp event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyChar As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Attribute KeyPress.VB_UserMemId = -603
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user moves the mouse into the control."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the user moves the mouse out of the control."
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
Private Declare PtrSafe Function SetLayout Lib "gdi32" (ByVal hDC As LongPtr, ByVal dwLayout As Long) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hWndFrom As LongPtr, ByVal hWndTo As LongPtr, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function GetMessagePos Lib "user32" () As Long
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal XY As Currency) As LongPtr
Private Declare PtrSafe Function GetTextAlign Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function SetTextAlign Lib "gdi32" (ByVal hDC As LongPtr, ByVal fMode As Long) As Long
Private Declare PtrSafe Function GetTextColor Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal crColor As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal XY As Currency) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal fMode As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
#End If
Private Const ICC_LINK_CLASS As Long = &H8000&
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
#If VBA7 Then
Private Const HWND_DESKTOP As LongPtr = &H0
#Else
Private Const HWND_DESKTOP As Long = &H0
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const LAYOUT_RTL As Long = &H1
Private Const TA_RTLREADING As Long = &H100
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINT As Long = &H317, PRF_CLIENT As Long = &H4, PRF_ERASEBKGND As Long = &H8
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const LWS_TRANSPARENT As Long = &H1 ' Unusable
Private Const LWS_IGNORERETURN As Long = &H2 ' Malfunction
Private Const LWS_NOPREFIX As Long = &H4
Private Const LWS_USEVISUALSTYLE As Long = &H8 ' Unusable
Private Const LWS_USECUSTOMTEXT As Long = &H10
Private Const LWS_RIGHT As Long = &H20
Private Const WM_USER As Long = &H400
Private Const LM_HITTEST As Long = (WM_USER + &H300)
Private Const LM_GETIDEALHEIGHT As Long = (WM_USER + &H301)
Private Const LM_GETIDEALSIZE As Long = LM_GETIDEALHEIGHT
Private Const LM_SETITEM As Long = (WM_USER + &H302)
Private Const LM_GETITEM As Long = (WM_USER + &H303)
Private Const LIF_ITEMINDEX As Long = &H1
Private Const LIF_STATE As Long = &H2
Private Const LIF_ITEMID As Long = &H4
Private Const LIF_URL As Long = &H8
Private Const LIS_FOCUSED As Long = &H1
Private Const LIS_ENABLED As Long = &H2
Private Const LIS_VISITED As Long = &H4
Private Const LIS_HOTTRACK As Long = &H8
Private Const LIS_DEFAULTCOLORS As Long = &H10
Private Const NM_FIRST As Long = 0
Private Const NM_CLICK As Long = (NM_FIRST - 2)
Private Const NM_RETURN As Long = (NM_FIRST - 4)
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ADDTOOL As Long = TTM_ADDTOOLW
#If VBA7 Then
Private Const LPSTR_TEXTCALLBACK As LongPtr = (-1)
#Else
Private Const LPSTR_TEXTCALLBACK As Long = (-1)
#End If
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_PARSELINKS As Long = &H1000
Private Const TTF_RTLREADING As Long = &H4
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Private LinkLabelHandle As LongPtr, LinkLabelToolTipHandle As LongPtr
Private LinkLabelTransparentBrush As LongPtr
Private LinkLabelFontHandle As LongPtr, LinkLabelUnderlineFontHandle As LongPtr
Private LinkLabelCharCodeCache As Long
Private LinkLabelMouseOver(0 To 3) As Boolean, LinkLabelMouseOverIndex As Long
Private LinkLabelDesignMode As Boolean
Private LinkLabelIsClick As Boolean
Private LinkLabelToolTipReady As Boolean
Private UCNoSetFocusFwd As Boolean

#If ImplementPreTranslateMsg = True Then

Private Const UM_PRETRANSLATEMSG As Long = (WM_USER + 1100)
Private UsePreTranslateMsg As Boolean

#End If

Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropLinks As LlbLinks
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBorderStyle As CCBorderStyleConstants
Private PropCaption As String
Private PropAlignment As CCLeftRightAlignmentConstants
Private PropHotTracking As Boolean
Private PropUnderlineHot As Boolean
Private PropUnderlineCold As Boolean
Private PropUseMnemonic As Boolean
Private PropTransparent As Boolean
Private PropShowTips As Boolean

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

#If VBA7 Then
Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal Shift As Long)
#Else
Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
#End If
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = CLng(wParam) And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyTab
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Dim Item As LITEM
            With Item
            .iLink = 0
            .Mask = LIF_ITEMINDEX Or LIF_STATE
            .StateMask = LIS_FOCUSED
            Do While SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0
                If .State = LIS_FOCUSED Then Handled = True
                .iLink = .iLink + 1
            Loop
            End With
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn, vbKeyEscape
            If IsInputKey = True Then
                SendMessage hWnd, wMsg, wParam, ByVal lParam
                Handled = True
            End If
    End Select
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_LINK_CLASS)

#If ImplementPreTranslateMsg = True Then

If SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject) = False Then UsePreTranslateMsg = True

#Else

Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)

#End If

End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
LinkLabelDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = CCBorderStyleNone
PropCaption = "<A>" & Ambient.DisplayName & "</A>"
If PropRightToLeft = False Then PropAlignment = CCLeftRightAlignmentLeft Else PropAlignment = CCLeftRightAlignmentRight
PropHotTracking = False
PropUnderlineHot = True
PropUnderlineCold = True
PropUseMnemonic = True
PropTransparent = False
PropShowTips = False
Call CreateLinkLabel
If LinkLabelHandle = NULL_PTR And ComCtlsSupportLevel() = 0 And LinkLabelDesignMode = True Then
    MsgBox "The LinkLabel control requires at least version 6.0 of comctl32.dll." & vbLf & _
    "In order to use it, you have to define a manifest file for your application." & vbLf & _
    "For using the control in the VB6 IDE, define a manifest file for VB6.EXE.", vbCritical + vbOKOnly
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
LinkLabelDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleNone)
PropCaption = VarToStr(.ReadProperty("Caption", vbNullString))
PropAlignment = .ReadProperty("Alignment", CCLeftRightAlignmentLeft)
PropHotTracking = .ReadProperty("HotTracking", False)
PropUnderlineHot = .ReadProperty("UnderlineHot", True)
PropUnderlineCold = .ReadProperty("UnderlineCold", True)
PropUseMnemonic = .ReadProperty("UseMnemonic", True)
PropTransparent = .ReadProperty("Transparent", False)
PropShowTips = .ReadProperty("ShowTips", False)
End With
Call CreateLinkLabel
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleNone
.WriteProperty "Caption", StrToVar(PropCaption), vbNullString
.WriteProperty "Alignment", PropAlignment, CCLeftRightAlignmentLeft
.WriteProperty "HotTracking", PropHotTracking, False
.WriteProperty "UnderlineHot", PropUnderlineHot, True
.WriteProperty "UnderlineCold", PropUnderlineCold, True
.WriteProperty "UseMnemonic", PropUseMnemonic, True
.WriteProperty "Transparent", PropTransparent, False
.WriteProperty "ShowTips", PropShowTips, False
End With
End Sub

Private Sub UserControl_Paint()
If LinkLabelHandle = NULL_PTR Then
    If UserControl.DrawStyle = vbInvisible Then UserControl.DrawStyle = vbSolid
    Dim i As Long
    For i = 8 To (UserControl.ScaleHeight + UserControl.ScaleWidth) Step 8
        UserControl.Line (-1, i)-(i, -1), vbBlack
    Next i
End If
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim P As POINTAPI
P.X = X
P.Y = Y
If LinkLabelHandle <> NULL_PTR Then MapWindowPoints UserControl.hWnd, LinkLabelHandle, P, 1
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim P As POINTAPI
P.X = X
P.Y = Y
If LinkLabelHandle <> NULL_PTR Then MapWindowPoints UserControl.hWnd, LinkLabelHandle, P, 1
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If LinkLabelHandle <> NULL_PTR Then
    If PropTransparent = True Then
        MoveWindow LinkLabelHandle, 0, 0, .ScaleWidth, .ScaleHeight, 0
        If LinkLabelTransparentBrush <> NULL_PTR Then
            DeleteObject LinkLabelTransparentBrush
            LinkLabelTransparentBrush = NULL_PTR
        End If
        RedrawWindow LinkLabelHandle, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
    Else
        MoveWindow LinkLabelHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    End If
    If PropShowTips = True And LinkLabelDesignMode = False Then
        Call DestroyToolTip
        Call CreateToolTip
    End If
End If
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()

#If ImplementPreTranslateMsg = True Then

If UsePreTranslateMsg = False Then Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)

#Else

Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)

#End If

Call DestroyLinkLabel
Call ComCtlsReleaseShellMod
End Sub

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
Attribute Container.VB_Description = "Returns the container of an object."
Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
Set Extender.Container = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
Extender.Top = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
Attribute ToolTipText.VB_MemberFlags = "400"
ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
Extender.ToolTipText = Value
End Property

Public Property Get HelpContextID() As Long
Attribute HelpContextID.VB_Description = "Specifies the default Help file context ID for an object."
HelpContextID = Extender.HelpContextID
End Property

Public Property Let HelpContextID(ByVal Value As Long)
Extender.HelpContextID = Value
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
Attribute WhatsThisHelpID.VB_MemberFlags = "400"
WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
Extender.WhatsThisHelpID = Value
End Property

Public Property Get DragIcon() As IPictureDisp
Attribute DragIcon.VB_Description = "Returns/sets the icon to be displayed as the pointer in a drag-and-drop operation."
Attribute DragIcon.VB_MemberFlags = "400"
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
Attribute DragMode.VB_Description = "Returns/sets a value that determines whether manual or automatic drag mode is used."
Attribute DragMode.VB_MemberFlags = "400"
DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
Attribute Drag.VB_Description = "Begins, ends, or cancels a drag operation of any object except Line, Menu, Shape, and Timer."
If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Moves the focus to the specified object."
Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
Attribute ZOrder.VB_Description = "Places a specified object at the front or back of the z-order within its graphical level."
If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

#If VBA7 Then
Public Property Get hWnd() As LongPtr
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#Else
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#End If
hWnd = LinkLabelHandle
End Property

#If VBA7 Then
Public Property Get hWndUserControl() As LongPtr
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
#Else
Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
#End If
hWndUserControl = UserControl.hWnd
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As LongPtr, OldUnderlineFontHandle As LongPtr
Dim TempFont As StdFont
Set PropFont = NewFont
OldFontHandle = LinkLabelFontHandle
OldUnderlineFontHandle = LinkLabelUnderlineFontHandle
LinkLabelFontHandle = CreateGDIFontFromOLEFont(PropFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Underline = True
LinkLabelUnderlineFontHandle = CreateGDIFontFromOLEFont(TempFont)
If LinkLabelHandle <> NULL_PTR Then
    SendMessage LinkLabelHandle, WM_SETFONT, LinkLabelFontHandle, ByVal 0&
    RedrawWindow LinkLabelHandle, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
End If
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
If OldUnderlineFontHandle <> NULL_PTR Then DeleteObject OldUnderlineFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As LongPtr, OldUnderlineFontHandle As LongPtr
Dim TempFont As StdFont
OldFontHandle = LinkLabelFontHandle
OldUnderlineFontHandle = LinkLabelUnderlineFontHandle
LinkLabelFontHandle = CreateGDIFontFromOLEFont(PropFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Underline = True
LinkLabelUnderlineFontHandle = CreateGDIFontFromOLEFont(TempFont)
If LinkLabelHandle <> NULL_PTR Then
    SendMessage LinkLabelHandle, WM_SETFONT, LinkLabelFontHandle, ByVal 0&
    RedrawWindow LinkLabelHandle, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
End If
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
If OldUnderlineFontHandle <> NULL_PTR Then DeleteObject OldUnderlineFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If LinkLabelHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles LinkLabelHandle
    Else
        RemoveVisualStyles LinkLabelHandle
    End If
    Call SetVisualStylesToolTip
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
UserControl.ForeColor = Value
Me.Refresh
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If LinkLabelHandle <> NULL_PTR Then EnableWindow LinkLabelHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDropMode() As OLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
Select Case Value
    Case OLEDropModeNone, OLEDropModeManual
        UserControl.OLEDropMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As CCMousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As CCMousePointerConstants)
Select Case Value
    Case 0 To 30, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
If LinkLabelDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = NULL_PTR Then
        Set PropMouseIcon = Value
    Else
        If LinkLabelDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If LinkLabelDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get MouseTrack() As Boolean
Attribute MouseTrack.VB_Description = "Returns/sets whether mouse events occurs when the mouse pointer enters or leaves the control."
MouseTrack = PropMouseTrack
End Property

Public Property Let MouseTrack(ByVal Value As Boolean)
PropMouseTrack = Value
UserControl.PropertyChanged "MouseTrack"
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system. This property is ignored at design time."
Attribute RightToLeft.VB_UserMemId = -611
RightToLeft = PropRightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
PropRightToLeft = Value
UserControl.RightToLeft = PropRightToLeft
Call ComCtlsCheckRightToLeft(PropRightToLeft, UserControl.RightToLeft, PropRightToLeftMode)
Dim dwMask As Long
If PropRightToLeft = True Then dwMask = WS_EX_RTLREADING Else dwMask = 0
If LinkLabelToolTipHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(LinkLabelToolTipHandle, dwMask)
Me.Refresh
UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftMode() As CCRightToLeftModeConstants
Attribute RightToLeftMode.VB_Description = "Returns/sets the right-to-left mode."
RightToLeftMode = PropRightToLeftMode
End Property

Public Property Let RightToLeftMode(ByVal Value As CCRightToLeftModeConstants)
Select Case Value
    Case CCRightToLeftModeNoControl, CCRightToLeftModeVBAME, CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
        PropRightToLeftMode = Value
    Case Else
        Err.Raise 380
End Select
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftMode"
End Property

Public Property Get BorderStyle() As CCBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As CCBorderStyleConstants)
Select Case Value
    Case CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If LinkLabelHandle <> NULL_PTR Then Call ComCtlsChangeBorderStyle(LinkLabelHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PPLinkLabelGeneral"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "121c"
If LinkLabelHandle <> NULL_PTR Then
    Caption = String$(CLng(SendMessage(LinkLabelHandle, WM_GETTEXTLENGTH, 0, ByVal 0&)), vbNullChar)
    SendMessage LinkLabelHandle, WM_GETTEXT, Len(Caption) + 1, ByVal StrPtr(Caption)
Else
    Caption = PropCaption
End If
End Property

Public Property Let Caption(ByVal Value As String)
PropCaption = Value
If LinkLabelHandle <> NULL_PTR Then
    SendMessage LinkLabelHandle, WM_SETTEXT, 0, ByVal StrPtr(PropCaption)
    If PropUseMnemonic = False And ComCtlsSupportLevel() >= 2 Then
        UserControl.AccessKeys = vbNullString
    Else
        UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
    End If
    Me.Refresh
    If PropShowTips = True And LinkLabelDesignMode = False Then
        Call DestroyToolTip
        Call CreateToolTip
    End If
End If
UserControl.PropertyChanged "Caption"
End Property

Public Property Get Alignment() As CCLeftRightAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment. Requires comctl32.dll version 6.1 or higher."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As CCLeftRightAlignmentConstants)
Select Case Value
    Case CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
If LinkLabelHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then Call ReCreateLinkLabel
UserControl.PropertyChanged "Alignment"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets whether hot tracking is enabled."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get UnderlineHot() As Boolean
Attribute UnderlineHot.VB_Description = "Returns/sets a value that determines whether hot link items to be displayed with underlined text or not. This property is ignored at design time."
UnderlineHot = PropUnderlineHot
End Property

Public Property Let UnderlineHot(ByVal Value As Boolean)
PropUnderlineHot = Value
Me.Refresh
UserControl.PropertyChanged "UnderlineHot"
End Property

Public Property Get UnderlineCold() As Boolean
Attribute UnderlineCold.VB_Description = "Returns/sets a value that determines whether cold link items to be displayed with underlined text or not. This property is ignored at design time."
UnderlineCold = PropUnderlineCold
End Property

Public Property Let UnderlineCold(ByVal Value As Boolean)
PropUnderlineCold = Value
Me.Refresh
UserControl.PropertyChanged "UnderlineCold"
End Property

Public Property Get UseMnemonic() As Boolean
Attribute UseMnemonic.VB_Description = "Returns/sets a value that specifies whether an & in the caption property defines an access key. Requires comctl32.dll version 6.1 or higher."
UseMnemonic = PropUseMnemonic
End Property

Public Property Let UseMnemonic(ByVal Value As Boolean)
PropUseMnemonic = Value
If LinkLabelHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then Call ReCreateLinkLabel
UserControl.PropertyChanged "UseMnemonic"
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value indicating if the background is a replica of the underlying background to simulate transparency. This property is ignored at design time."
Transparent = PropTransparent
End Property

Public Property Let Transparent(ByVal Value As Boolean)
PropTransparent = Value
Me.Refresh
UserControl.PropertyChanged "Transparent"
End Property

Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Returns/sets a value that determines whether the LinkGetTipText event will be raised to retrieve a tool tip text to be displayed or not."
ShowTips = PropShowTips
End Property

Public Property Let ShowTips(ByVal Value As Boolean)
PropShowTips = Value
If LinkLabelHandle <> NULL_PTR And LinkLabelDesignMode = False Then
    Call DestroyToolTip
    If PropShowTips = True Then Call CreateToolTip
End If
UserControl.PropertyChanged "ShowTips"
End Property

Public Property Get Links() As LlbLinks
Attribute Links.VB_Description = "Returns a reference to a collection of the link objects."
If PropLinks Is Nothing Then
    Set PropLinks = New LlbLinks
    PropLinks.FInit Me
End If
Set Links = PropLinks
End Property

Friend Function FLinksCount() As Long
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = 0
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    Do While SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0
        .iLink = .iLink + 1
    Loop
    FLinksCount = .iLink
    End With
End If
End Function

Friend Property Get FLinkURL(ByVal Index As Long) As String
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_URL
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        FLinkURL = Left$(.szURL, InStr(.szURL, vbNullChar) - 1)
    End If
    End With
End If
End Property

Friend Property Let FLinkURL(ByVal Index As Long, ByVal Value As String)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM, Buffer As String
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_URL
    Buffer = Left$(Value, L_MAX_URL_LENGTH)
    CopyMemory .szURL(0), ByVal StrPtr(Buffer), LenB(Buffer)
    End With
    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
End If
End Property

Friend Property Get FLinkIDName(ByVal Index As Long) As String
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_ITEMID
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        FLinkIDName = Left$(.szID, InStr(.szID, vbNullChar) - 1)
    End If
    End With
End If
End Property

Friend Property Let FLinkIDName(ByVal Index As Long, ByVal Value As String)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM, Buffer As String
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_ITEMID
    Buffer = Left$(Value, MAX_LINKID_TEXT)
    CopyMemory .szID(0), ByVal StrPtr(Buffer), LenB(Buffer)
    End With
    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
End If
End Property

Friend Property Get FLinkCaption(ByVal Index As Long) As String
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        Dim Temp As String, j As Long, i As Long, CharPosEnd As Long, CharPosStart As Long
        Temp = Me.Caption
        j = 1
        Do While Index > 0
            CharPosEnd = InStr(j, Temp, "</A>", vbTextCompare)
            For i = CharPosEnd To 2 Step -1
                If StrComp(Mid$(Temp, i - 1, 1), ">", vbTextCompare) = 0 Then
                    CharPosStart = i
                    Index = Index - 1
                    Exit For
                End If
            Next i
            j = CharPosEnd + 4 ' Len of "</A>"
        Loop
        If CharPosStart > 0 And CharPosEnd > 0 Then FLinkCaption = Mid$(Temp, CharPosStart, CharPosEnd - CharPosStart)
    End If
    End With
End If
End Property

Friend Property Let FLinkCaption(ByVal Index As Long, ByVal Value As String)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        Dim Temp As String, j As Long, i As Long, CharPosEnd As Long, CharPosStart As Long
        Temp = Me.Caption
        j = 1
        Do While Index > 0
            CharPosEnd = InStr(j, Temp, "</A>", vbTextCompare)
            For i = CharPosEnd To 2 Step -1
                If StrComp(Mid$(Temp, i - 1, 1), ">", vbTextCompare) = 0 Then
                    CharPosStart = i
                    Index = Index - 1
                    Exit For
                End If
            Next i
            j = CharPosEnd + 4 ' Len of "</A>"
        Loop
        If CharPosStart > 0 And CharPosEnd > 0 Then
            Dim Text As String
            Text = Left$(Temp, CharPosStart - 1) & Value & Mid$(Temp, CharPosEnd)
            Me.Caption = Text
        End If
    End If
    End With
End If
End Property

Friend Property Get FLinkSelected(ByVal Index As Long) As Boolean
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_FOCUSED
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        FLinkSelected = CBool(.State = LIS_FOCUSED)
    End If
    End With
End If
End Property

Friend Property Let FLinkSelected(ByVal Index As Long, ByVal Value As Boolean)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_FOCUSED
    If Value = True Then
        .State = LIS_FOCUSED
    Else
        .State = 0
    End If
    End With
    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
End If
End Property

Friend Property Get FLinkEnabled(ByVal Index As Long) As Boolean
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_ENABLED
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        FLinkEnabled = CBool(.State = LIS_ENABLED)
    End If
    End With
End If
End Property

Friend Property Let FLinkEnabled(ByVal Index As Long, ByVal Value As Boolean)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_ENABLED
    If Value = True Then
        .State = LIS_ENABLED
    Else
        .State = 0
    End If
    End With
    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
End If
End Property

Friend Property Get FLinkVisited(ByVal Index As Long) As Boolean
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_VISITED
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        FLinkVisited = CBool(.State = LIS_VISITED)
    End If
    End With
End If
End Property

Friend Property Let FLinkVisited(ByVal Index As Long, ByVal Value As Boolean)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_VISITED
    If Value = True Then
        .State = LIS_VISITED
    Else
        .State = 0
    End If
    End With
    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
End If
End Property

Friend Property Get FLinkHot(ByVal Index As Long) As Boolean
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_HOTTRACK
    If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
        FLinkHot = CBool(.State = LIS_HOTTRACK)
    End If
    End With
End If
End Property

Friend Property Let FLinkHot(ByVal Index As Long, ByVal Value As Boolean)
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = Index - 1
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_HOTTRACK
    If Value = True Then
        .State = LIS_HOTTRACK
    Else
        .State = 0
    End If
    End With
    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
End If
End Property

Friend Property Get FLinkLeft(ByVal Index As Long) As Single
If LinkLabelHandle <> NULL_PTR Then
    Dim RC As RECT
    Call GetLinkRect(Index, RC)
    FLinkLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FLinkTop(ByVal Index As Long) As Single
If LinkLabelHandle <> NULL_PTR Then
    Dim RC As RECT
    Call GetLinkRect(Index, RC)
    FLinkTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FLinkWidth(ByVal Index As Long) As Single
If LinkLabelHandle <> NULL_PTR Then
    Dim RC As RECT
    Call GetLinkRect(Index, RC)
    FLinkWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FLinkHeight(ByVal Index As Long) As Single
If LinkLabelHandle <> NULL_PTR Then
    Dim RC As RECT
    Call GetLinkRect(Index, RC)
    FLinkHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Private Sub CreateLinkLabel()
If LinkLabelHandle <> NULL_PTR Or ComCtlsSupportLevel() = 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP
' According to MSDN:
' Bidirectional display (WS_EX_RTLREADING) is not supported directly.
' However, this can be achieved when using SetTextAlign TA_RTLREADING on CDDS_PREPAINT in NM_CUSTOMDRAW.
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
If PropAlignment = CCLeftRightAlignmentRight And ComCtlsSupportLevel() >= 2 Then dwStyle = dwStyle Or LWS_RIGHT
If PropUseMnemonic = False And ComCtlsSupportLevel() >= 2 Then dwStyle = dwStyle Or LWS_NOPREFIX
LinkLabelHandle = CreateWindowEx(dwExStyle, StrPtr("SysLink"), NULL_PTR, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Caption = PropCaption
If LinkLabelDesignMode = False Then
    If LinkLabelHandle <> NULL_PTR Then Call ComCtlsSetSubclass(LinkLabelHandle, Me, 1)
    If LinkLabelHandle <> NULL_PTR Then
        ' This trick allows the usage of the GetLinkRect method at initialization time.
        ' Must be called after WM_SETTEXT is processed.
        ' Only after this the 'ShowTips' property can be set.
        SendMessage LinkLabelHandle, WM_PAINT, 0, ByVal 0&
        LinkLabelToolTipReady = True
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
    
    #If ImplementPreTranslateMsg = True Then
    
    If UsePreTranslateMsg = True Then Call ComCtlsPreTranslateMsgAddHook
    
    #End If
    
End If
Me.ShowTips = PropShowTips
End Sub

Private Sub CreateToolTip()
Static Done As Boolean
Dim Count As Long, dwExStyle As Long
Count = Me.FLinksCount
If LinkLabelToolTipHandle <> NULL_PTR Or Count = 0 Or LinkLabelToolTipReady = False Then Exit Sub
If Done = False Then
    Call ComCtlsInitCC(ICC_TAB_CLASSES)
    Done = True
End If
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
LinkLabelToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If LinkLabelToolTipHandle <> NULL_PTR Then Call ComCtlsInitToolTip(LinkLabelToolTipHandle)
Call SetVisualStylesToolTip
If LinkLabelToolTipHandle <> NULL_PTR Then
    Dim TI As TOOLINFO, i As Long
    With TI
    .cbSize = LenB(TI)
    .hWnd = LinkLabelHandle
    For i = 1 To Count
        .uId = i
        .uFlags = TTF_SUBCLASS Or TTF_PARSELINKS
        If PropRightToLeft = True Then .uFlags = .uFlags Or TTF_RTLREADING
        .lpszText = LPSTR_TEXTCALLBACK
        Call GetLinkRect(i, .RC)
        SendMessage LinkLabelToolTipHandle, TTM_ADDTOOL, 0, ByVal VarPtr(TI)
    Next i
    End With
End If
End Sub

Private Sub ReCreateLinkLabel()
If LinkLabelDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Call DestroyLinkLabel
    Call CreateLinkLabel
    Call UserControl_Resize
    If Locked = True Then LockWindowUpdate NULL_PTR
    Me.Refresh
Else
    Call DestroyLinkLabel
    Call CreateLinkLabel
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyLinkLabel()
If LinkLabelHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(LinkLabelHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
If LinkLabelDesignMode = False Then
    
    #If ImplementPreTranslateMsg = True Then
    
    If UsePreTranslateMsg = True Then Call ComCtlsPreTranslateMsgReleaseHook
    
    #End If
    
End If
Call DestroyToolTip
ShowWindow LinkLabelHandle, SW_HIDE
SetParent LinkLabelHandle, NULL_PTR
DestroyWindow LinkLabelHandle
LinkLabelHandle = NULL_PTR
If LinkLabelFontHandle <> NULL_PTR Then
    DeleteObject LinkLabelFontHandle
    LinkLabelFontHandle = NULL_PTR
End If
If LinkLabelUnderlineFontHandle <> NULL_PTR Then
    DeleteObject LinkLabelUnderlineFontHandle
    LinkLabelUnderlineFontHandle = NULL_PTR
End If
If LinkLabelTransparentBrush <> NULL_PTR Then
    DeleteObject LinkLabelTransparentBrush
    LinkLabelTransparentBrush = NULL_PTR
End If
End Sub

Private Sub DestroyToolTip()
If LinkLabelToolTipHandle = NULL_PTR Then Exit Sub
DestroyWindow LinkLabelToolTipHandle
LinkLabelToolTipHandle = NULL_PTR
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If LinkLabelTransparentBrush <> NULL_PTR Then
    DeleteObject LinkLabelTransparentBrush
    LinkLabelTransparentBrush = NULL_PTR
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As LlbLink
Attribute HitTest.VB_Description = "Returns a reference to the link item object located at the coordinates of X and Y."
If LinkLabelHandle <> NULL_PTR Then
    Dim LHTI As LHITTESTINFO
    With LHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI)) <> 0 Then
        Set HitTest = Me.Links(.Item.iLink + 1)
    End If
    End With
End If
End Function

Public Property Get SelectedItem() As LlbLink
Attribute SelectedItem.VB_Description = "Returns/sets a reference to the currently selected link item."
Attribute SelectedItem.VB_MemberFlags = "400"
If LinkLabelHandle <> NULL_PTR Then
    Dim Item As LITEM
    With Item
    .iLink = 0
    .Mask = LIF_ITEMINDEX Or LIF_STATE
    .StateMask = LIS_FOCUSED
    Do While SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0
        .iLink = .iLink + 1
        If .State = LIS_FOCUSED Then
            Set SelectedItem = Me.Links(.iLink)
            Exit Do
        End If
    Loop
    End With
End If
End Property

Public Property Let SelectedItem(ByVal Value As LlbLink)
Set Me.SelectedItem = Value
End Property

Public Property Set SelectedItem(ByVal Value As LlbLink)
If LinkLabelHandle <> NULL_PTR Then
    If Not Value Is Nothing Then
        Value.Selected = True
    Else
        Dim Item As LITEM
        With Item
        .iLink = 0
        .Mask = LIF_ITEMINDEX Or LIF_STATE
        .StateMask = LIS_FOCUSED
        Do While SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0
            If .State = LIS_FOCUSED Then
                .State = 0
                SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
            End If
            .iLink = .iLink + 1
        Loop
        End With
    End If
End If
End Property

Public Function GetIdealHeight() As Single
Attribute GetIdealHeight.VB_Description = "Gets the ideal height of the control."
If LinkLabelHandle <> NULL_PTR Then
    Dim RC(0 To 1) As RECT, RetVal As Long
    GetWindowRect LinkLabelHandle, RC(0)
    GetClientRect LinkLabelHandle, RC(1)
    RetVal = CLng(SendMessage(LinkLabelHandle, LM_GETIDEALHEIGHT, 0, ByVal 0&))
    RetVal = RetVal + ((RC(0).Bottom - RC(0).Top) - (RC(1).Bottom - RC(1).Top))
    With UserControl
    GetIdealHeight = .ScaleY(RetVal, vbPixels, vbContainerSize)
    End With
End If
End Function

Public Sub GetIdealSize(ByRef Width As Single, ByRef Height As Single)
Attribute GetIdealSize.VB_Description = "Gets the ideal size of the control. Requires comctl32.dll version 6.1 or higher."
If LinkLabelHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then
    Width = 0
    Height = 0
    Dim RC(0 To 1) As RECT
    GetWindowRect LinkLabelHandle, RC(0)
    GetClientRect LinkLabelHandle, RC(1)
    Dim Size As SIZEAPI
    SendMessage LinkLabelHandle, LM_GETIDEALSIZE, RC(1).Right - RC(1).Left, ByVal VarPtr(Size)
    Size.CX = Size.CX + ((RC(0).Right - RC(0).Left) - (RC(1).Right - RC(1).Left))
    Size.CY = Size.CY + ((RC(0).Bottom - RC(0).Top) - (RC(1).Bottom - RC(1).Top))
    With UserControl
    Width = .ScaleX(Size.CX, vbPixels, vbContainerSize)
    Height = .ScaleY(Size.CY, vbPixels, vbContainerSize)
    End With
End If
End Sub

Private Sub GetLinkRect(ByVal Index As Long, ByRef RC As RECT)
If LinkLabelHandle <> NULL_PTR Then
    Dim LHTI As LHITTESTINFO, ClientRect As RECT, Success As Boolean
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    GetClientRect LinkLabelHandle, ClientRect
    With LHTI
    For X1 = ClientRect.Left To ClientRect.Right
        For Y1 = ClientRect.Top To ClientRect.Bottom
            .PT.X = X1
            .PT.Y = Y1
            If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI)) <> 0 Then
                If .Item.iLink = Index - 1 Then
                    Success = True
                    Exit For
                End If
            End If
        Next Y1
        If Success = True Then Exit For
    Next X1
    If Success = True Then
        For X2 = (X1 + 1) To ClientRect.Right
            .PT.X = X2
            .PT.Y = Y1
            If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI)) <> 0 Then
                If .Item.iLink <> Index - 1 Then Exit For
            Else
                Exit For
            End If
        Next X2
        For Y2 = (Y1 + 1) To ClientRect.Bottom
            .PT.X = X1
            .PT.Y = Y2
            If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI)) <> 0 Then
                If .Item.iLink <> Index - 1 Then Exit For
            Else
                Exit For
            End If
        Next Y2
        RC.Left = X1
        RC.Right = X2
        RC.Top = Y1
        RC.Bottom = Y2
    End If
    End With
End If
End Sub

Private Sub SetVisualStylesToolTip()
If LinkLabelHandle <> NULL_PTR Then
    If LinkLabelToolTipHandle <> NULL_PTR And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles LinkLabelToolTipHandle
        Else
            RemoveVisualStyles LinkLabelToolTipHandle
        End If
    End If
End If
End Sub

#If ImplementPreTranslateMsg = True Then

Private Function PreTranslateMsg(ByVal lParam As LongPtr) As LongPtr
PreTranslateMsg = 0
If lParam <> NULL_PTR Then
    Dim Msg As TMSG, Handled As Boolean, RetVal As Long
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    IOleInPlaceActiveObjectVB_TranslateAccelerator Handled, RetVal, Msg.hWnd, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
    If Handled = True Then PreTranslateMsg = 1
End If
End Function

#End If

#If VBA7 Then
Private Function ISubclass_Message(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
#End If
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        
        #If ImplementPreTranslateMsg = True Then
        
        If UsePreTranslateMsg = False Then Call ActivateIPAO(Me) Else Call ComCtlsPreTranslateMsgActivate(hWnd)
        
        #Else
        
        Call ActivateIPAO(Me)
        
        #End If
        
    Case WM_KILLFOCUS
        
        #If ImplementPreTranslateMsg = True Then
        
        If UsePreTranslateMsg = False Then Call DeActivateIPAO Else Call ComCtlsPreTranslateMsgDeActivate
        
        #Else
        
        Call DeActivateIPAO
        
        #End If
        
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = CLng(wParam) And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            LinkLabelCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If LinkLabelCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(LinkLabelCharCodeCache And &HFFFF&)
            LinkLabelCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(CLng(wParam) And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcControl = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(CLng(wParam))
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcControl = 0
        End If
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    Case WM_MOUSEMOVE
        If PropHotTracking = True Then
            Dim LHTI1 As LHITTESTINFO, Index As Long
            With LHTI1
            .PT.X = Get_X_lParam(lParam)
            .PT.Y = Get_Y_lParam(lParam)
            With .Item
            If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI1)) <> 0 Then
                Index = .iLink + 1
            End If
            .iLink = 0
            .Mask = LIF_ITEMINDEX Or LIF_STATE
            .StateMask = LIS_HOTTRACK
            Do While SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(LHTI1.Item)) <> 0
                If .State = LIS_HOTTRACK Then
                    If .iLink <> Index - 1 Or Index = 0 Then
                        .State = 0
                        SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(LHTI1.Item)
                    End If
                Else
                    If .iLink = Index - 1 And Index > 0 Then
                        .State = LIS_HOTTRACK
                        SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(LHTI1.Item)
                    End If
                End If
                .iLink = .iLink + 1
            Loop
            End With
            End With
        End If
    Case WM_MOUSELEAVE
        If PropHotTracking = True Then
            Dim Item As LITEM
            With Item
            .iLink = 0
            .Mask = LIF_ITEMINDEX Or LIF_STATE
            .StateMask = LIS_HOTTRACK
            Do While SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0
                If .State = LIS_HOTTRACK Then
                    .State = 0
                    SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
                End If
                .iLink = .iLink + 1
            Loop
            End With
        End If
    Case WM_CONTEXTMENU
        If wParam = LinkLabelHandle Then
            Dim P1 As POINTAPI
            P1.X = Get_X_lParam(lParam)
            P1.Y = Get_Y_lParam(lParam)
            If P1.X = -1 And P1.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(-1, -1)
            Else
                ScreenToClient LinkLabelHandle, P1
                RaiseEvent ContextMenu(UserControl.ScaleX(P1.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P1.Y, vbPixels, vbContainerPosition))
            End If
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = LinkLabelToolTipHandle And LinkLabelToolTipHandle <> NULL_PTR Then
            Select Case NM.Code
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    With NMTTDI
                    Dim Text As String
                    RaiseEvent LinkGetTipText(Me.Links(CLng(NM.IDFrom)), Text)
                    If Not Text = vbNullString Then
                        If Len(Text) <= 80 Then
                            Text = Left$(Text & vbNullChar, 80)
                            CopyMemory .szText(0), ByVal StrPtr(Text), LenB(Text)
                        Else
                            .lpszText = StrPtr(Text)
                        End If
                        .hInst = NULL_PTR
                        CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                    End If
                    End With
            End Select
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If wParam = LinkLabelToolTipHandle And LinkLabelToolTipHandle <> NULL_PTR And lParam = NF_QUERY Then
            Const NFR_ANSI As Long = 1
            Const NFR_UNICODE As Long = 2
            WindowProcControl = NFR_UNICODE
            Exit Function
        End If
    Case WM_PAINT
        If wParam <> 0 Then
            SendMessage hWnd, WM_PRINT, wParam, ByVal PRF_CLIENT Or PRF_ERASEBKGND
            WindowProcControl = 0
            Exit Function
        End If
    
    #If ImplementPreTranslateMsg = True Then
    
    Case UM_PRETRANSLATEMSG
        WindowProcControl = PreTranslateMsg(lParam)
        Exit Function
    
    #End If
    
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                LinkLabelIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                LinkLabelIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                LinkLabelIsClick = True
            Case WM_MOUSEMOVE
                If (LinkLabelMouseOver(0) = False And PropHotTracking = True) Or (LinkLabelMouseOver(1) = False And PropMouseTrack = True) Or (LinkLabelMouseOver(3) = False And PropMouseTrack = True) Then
                    If LinkLabelMouseOver(0) = False And PropHotTracking = True Then LinkLabelMouseOver(0) = True
                    If LinkLabelMouseOver(1) = False And PropMouseTrack = True Then LinkLabelMouseOver(1) = True
                    If LinkLabelMouseOver(3) = False And PropMouseTrack = True Then
                        LinkLabelMouseOver(3) = True
                        RaiseEvent MouseEnter
                    End If
                    If LinkLabelMouseOver(1) = True And PropMouseTrack = True Then
                        Dim LHTI2 As LHITTESTINFO
                        With LHTI2
                        .PT.X = Get_X_lParam(lParam)
                        .PT.Y = Get_Y_lParam(lParam)
                        If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI2)) <> 0 Then
                            LinkLabelMouseOverIndex = .Item.iLink + 1
                        Else
                            LinkLabelMouseOverIndex = 0
                        End If
                        End With
                        If LinkLabelMouseOverIndex > 0 Then RaiseEvent LinkMouseEnter(Me.Links(LinkLabelMouseOverIndex))
                    End If
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                If LinkLabelMouseOver(1) = True And PropMouseTrack = True Then
                    Dim LHTI3 As LHITTESTINFO
                    With LHTI3
                    .PT.X = Get_X_lParam(lParam)
                    .PT.Y = Get_Y_lParam(lParam)
                    If SendMessage(LinkLabelHandle, LM_HITTEST, 0, ByVal VarPtr(LHTI3)) <> 0 Then
                        If LinkLabelMouseOverIndex <> .Item.iLink + 1 Then
                            If LinkLabelMouseOverIndex > 0 Then RaiseEvent LinkMouseLeave(Me.Links(LinkLabelMouseOverIndex))
                            LinkLabelMouseOverIndex = .Item.iLink + 1
                            If LinkLabelMouseOverIndex > 0 Then RaiseEvent LinkMouseEnter(Me.Links(LinkLabelMouseOverIndex))
                        End If
                    Else
                        If LinkLabelMouseOverIndex > 0 Then RaiseEvent LinkMouseLeave(Me.Links(LinkLabelMouseOverIndex))
                        LinkLabelMouseOverIndex = 0
                    End If
                    End With
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                End Select
                If LinkLabelIsClick = True Then
                    LinkLabelIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        LinkLabelMouseOver(0) = False
        If LinkLabelMouseOver(1) = True Then
            LinkLabelMouseOver(1) = False
            If LinkLabelMouseOverIndex > 0 Then RaiseEvent LinkMouseLeave(Me.Links(LinkLabelMouseOverIndex))
        End If
        If LinkLabelMouseOver(3) = True Then
            Dim Pos As Long, P2 As POINTAPI, XY As Currency
            Pos = GetMessagePos()
            P2.X = Get_X_lParam(Pos)
            P2.Y = Get_Y_lParam(Pos)
            CopyMemory ByVal VarPtr(XY), ByVal VarPtr(P2), 8
            If WindowFromPoint(XY) <> UserControl.hWnd Then
                LinkLabelMouseOver(3) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = LinkLabelHandle Then
            Select Case NM.Code
                Case NM_CUSTOMDRAW
                    Dim NMCD As NMCUSTOMDRAW
                    CopyMemory NMCD, ByVal lParam, LenB(NMCD)
                    Select Case NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            If PropRightToLeft = True Then
                                Dim fMode As Long
                                fMode = GetTextAlign(NMCD.hDC)
                                If (fMode And TA_RTLREADING) = 0 Then fMode = fMode Or TA_RTLREADING
                                SetTextAlign NMCD.hDC, fMode
                            End If
                            WindowProcUserControl = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            If NMCD.dwItemSpec > -1 Then
                                Dim FontHandle As LongPtr, RGBColor As Long, OldColor As Long
                                FontHandle = LinkLabelFontHandle
                                If Me.FLinkHot(CLng(NMCD.dwItemSpec) + 1) = False Then
                                    If PropUnderlineCold = True Then FontHandle = LinkLabelUnderlineFontHandle
                                Else
                                    If PropUnderlineHot = True Then FontHandle = LinkLabelUnderlineFontHandle
                                End If
                                SelectObject NMCD.hDC, FontHandle
                                If ComCtlsSupportLevel >= 2 Then
                                    RGBColor = GetTextColor(NMCD.hDC)
                                    OldColor = RGBColor
                                    RaiseEvent LinkForeColor(Me.Links(CLng(NMCD.dwItemSpec) + 1), RGBColor)
                                    If OldColor <> RGBColor Then SetTextColor NMCD.hDC, RGBColor
                                End If
                                WindowProcUserControl = CDRF_NEWFONT
                            Else
                                WindowProcUserControl = CDRF_DODEFAULT
                            End If
                            Exit Function
                    End Select
                Case NM_CLICK, NM_RETURN
                    Dim NML As NMLINK, Reason As LlbLinkActivateReasonConstants
                    CopyMemory NML, ByVal lParam, LenB(NML)
                    If NML.Item.iLink > -1 Then
                        Select Case NM.Code
                            Case NM_CLICK
                                Reason = LlbLinkActivateReasonClick
                            Case NM_RETURN
                                Reason = LlbLinkActivateReasonReturn
                        End Select
                        Dim Item As LITEM, PrevState As Long
                        With Item
                        .iLink = NML.Item.iLink
                        .Mask = LIF_ITEMINDEX Or LIF_STATE
                        .StateMask = LIS_FOCUSED
                        If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then PrevState = .State
                        RaiseEvent LinkActivate(Me.Links(NML.Item.iLink + 1), Reason)
                        Select Case GetFocus()
                            Case UserControl.hWnd, LinkLabelHandle
                                If SendMessage(LinkLabelHandle, LM_GETITEM, 0, ByVal VarPtr(Item)) <> 0 Then
                                    If .State = 0 And PrevState = LIS_FOCUSED Then
                                        .State = LIS_FOCUSED
                                        SendMessage LinkLabelHandle, LM_SETITEM, 0, ByVal VarPtr(Item)
                                    End If
                                End If
                        End Select
                        End With
                    End If
            End Select
        End If
    Case WM_SETCURSOR
        If LoWord(CLng(lParam)) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(NULL_PTR, MousePointerID(PropMousePointer))
                WindowProcUserControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcUserControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_CTLCOLORSTATIC
        WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropTransparent = True Then
            SetBkMode wParam, 1
            Dim hDCBmp As LongPtr
            Dim hBmp As LongPtr, hBmpOld As LongPtr
            With UserControl
            If LinkLabelTransparentBrush = NULL_PTR Then
                hDCBmp = CreateCompatibleDC(wParam)
                If hDCBmp <> NULL_PTR Then
                    hBmp = CreateCompatibleBitmap(wParam, .ScaleWidth, .ScaleHeight)
                    If hBmp <> NULL_PTR Then
                        Dim hWndParent As LongPtr
                        hWndParent = GetParent(.hWnd)
                        If (GetWindowLong(hWndParent, GWL_EXSTYLE) And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then SetLayout hDCBmp, LAYOUT_RTL
                        hBmpOld = SelectObject(hDCBmp, hBmp)
                        Dim WndRect As RECT, P1 As POINTAPI
                        GetWindowRect .hWnd, WndRect
                        MapWindowPoints HWND_DESKTOP, hWndParent, WndRect, 2
                        P1.X = WndRect.Left
                        P1.Y = WndRect.Top
                        SetViewportOrgEx hDCBmp, -P1.X, -P1.Y, P1
                        SendMessage hWndParent, WM_PAINT, hDCBmp, ByVal 0&
                        SetViewportOrgEx hDCBmp, P1.X, P1.Y, P1
                        LinkLabelTransparentBrush = CreatePatternBrush(hBmp)
                        SelectObject hDCBmp, hBmpOld
                        DeleteObject hBmp
                    End If
                    DeleteDC hDCBmp
                End If
            End If
            End With
            If LinkLabelTransparentBrush <> NULL_PTR Then WindowProcUserControl = LinkLabelTransparentBrush
        End If
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_SETFOCUS
        If UCNoSetFocusFwd = False Then SetFocusAPI LinkLabelHandle
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                LinkLabelIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                LinkLabelIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                LinkLabelIsClick = True
            Case WM_MOUSEMOVE
                If (LinkLabelMouseOver(2) = False And PropMouseTrack = True) Or (LinkLabelMouseOver(3) = False And PropMouseTrack = True) Then
                    If LinkLabelMouseOver(2) = False And PropMouseTrack = True Then LinkLabelMouseOver(2) = True
                    If LinkLabelMouseOver(3) = False And PropMouseTrack = True Then
                        LinkLabelMouseOver(3) = True
                        RaiseEvent MouseEnter
                    End If
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                End Select
                If LinkLabelIsClick = True Then
                    LinkLabelIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        LinkLabelMouseOver(2) = False
        If LinkLabelMouseOver(3) = True Then
            Dim Pos As Long, P2 As POINTAPI, XY As Currency
            Pos = GetMessagePos()
            P2.X = Get_X_lParam(Pos)
            P2.Y = Get_Y_lParam(Pos)
            CopyMemory ByVal VarPtr(XY), ByVal VarPtr(P2), 8
            If WindowFromPoint(XY) <> LinkLabelHandle Or LinkLabelHandle = NULL_PTR Then
                LinkLabelMouseOver(3) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function
