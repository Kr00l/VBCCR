VERSION 5.00
Begin VB.UserControl TabStrip 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "TabStrip.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "TabStrip.ctx":005A
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "TabStrip"
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
Private TbsPlacementTop, TbsPlacementBottom, TbsPlacementLeft, TbsPlacementRight
Private TbsStyleTabs, TbsStyleButtons, TbsStyleFlatButtons
Private TbsTabStyleStandard, TbsTabStyleOpposite
Private TbsTabWidthStyleJustified, TbsTabWidthStyleNonJustified, TbsTabWidthStyleFixed
Private TbsTabAlignmentStandard, TbsTabAlignmentImageLeft, TbsTabAlignmentImageCaptionLeft
Private TbsHitResultNoWhere, TbsHitResultItem, TbsHitResultItemIcon, TbsHitResultItemLabel
Private TbsDrawModeNormal, TbsDrawModeOwnerDrawFixed
#End If
Public Enum TbsPlacementConstants
TbsPlacementTop = 0
TbsPlacementBottom = 1
TbsPlacementLeft = 2
TbsPlacementRight = 3
End Enum
Public Enum TbsStyleConstants
TbsStyleTabs = 0
TbsStyleButtons = 1
TbsStyleFlatButtons = 2
End Enum
Public Enum TbsTabStyleConstants
TbsTabStyleStandard = 0
TbsTabStyleOpposite = 1
End Enum
Public Enum TbsTabWidthStyleConstants
TbsTabWidthStyleJustified = 0
TbsTabWidthStyleNonJustified = 1
TbsTabWidthStyleFixed = 2
End Enum
Public Enum TbsTabAlignmentConstants
TbsTabAlignmentStandard = 0
TbsTabAlignmentImageLeft = 1
TbsTabAlignmentImageCaptionLeft = 2
End Enum
Public Enum TbsHitResultConstants
TbsHitResultNoWhere = 0
TbsHitResultItem = 1
TbsHitResultItemIcon = 2
TbsHitResultItemLabel = 3
End Enum
Public Enum TbsDrawModeConstants
TbsDrawModeNormal = 0
TbsDrawModeOwnerDrawFixed = 1
End Enum
Private Type TACCEL
FVirt As Byte
Key As Integer
Cmd As Integer
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
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
Private Type TCITEM
Mask As Long
dwState As Long
dwStateMask As Long
pszText As LongPtr
cchTextMax As Long
iImage As Long
lParam As LongPtr
End Type
Private Type TCHITTESTINFO
PT As POINTAPI
Flags As Long
End Type
Private Type NMHDR
hWndFrom As LongPtr
IDFrom As LongPtr
Code As Long
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As LongPtr
szText(0 To ((80 * 2) - 1)) As Byte
hInst As LongPtr
uFlags As Long
lParam As LongPtr
End Type
Private Type PAINTSTRUCT
hDC As LongPtr
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
End Type
Private Type DRAWITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemAction As Long
ItemState As Long
hWndItem As LongPtr
hDC As LongPtr
RCItem As RECT
ItemData As LongPtr
End Type
Public Event TabBeforeClick(ByVal TabItem As TbsTab, ByRef Cancel As Boolean)
Attribute TabBeforeClick.VB_Description = "Occurs when a tab is clicked, or the tab's value setting has been changed. Used to check parameters before actually generating a TabClick event."
Public Event TabClick(ByVal TabItem As TbsTab)
Attribute TabClick.VB_Description = "Occurs when a tab is clicked, or the tab's value setting has been changed."
Public Event ItemDraw(ByVal TabItem As TbsTab, ByVal ItemAction As Long, ByVal ItemState As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Attribute ItemDraw.VB_Description = "Occurs when a visual aspect of an owner-drawn tab strip has changed."
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
Private Declare PtrSafe Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableW" (ByVal lpAccel As LongPtr, ByVal cEntries As Long) As LongPtr
Private Declare PtrSafe Function DestroyAcceleratorTable Lib "user32" (ByVal hAccel As LongPtr) As Long
Private Declare PtrSafe Function VkKeyScan Lib "user32" Alias "VkKeyScanW" (ByVal cChar As Integer) As Integer
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hWndFrom As LongPtr, ByVal hWndTo As LongPtr, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPaint As PAINTSTRUCT) As LongPtr
Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare PtrSafe Function WindowFromDC Lib "user32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, ByRef lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
Private Declare PtrSafe Function CombineRgn Lib "gdi32" (ByVal hRgnDest As LongPtr, ByVal hRgnSrc1 As LongPtr, ByVal hRgnSrc2 As LongPtr, ByVal nCombineMode As Long) As Long
Private Declare PtrSafe Function FillRgn Lib "gdi32" (ByVal hDC As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableW" (ByVal lpAccel As Long, ByVal cEntries As Long) As Long
Private Declare Function DestroyAcceleratorTable Lib "user32" (ByVal hAccel As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanW" (ByVal cChar As Integer) As Integer
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hRgnDest As Long, ByVal hRgnSrc1 As Long, ByVal hRgnSrc2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
#If VBA7 Then
Private Const HWND_DESKTOP As LongPtr = &H0
#Else
Private Const HWND_DESKTOP As Long = &H0
#End If
Private Const COLOR_BTNFACE As Long = 15
Private Const RGN_OR As Long = 2
Private Const RGN_DIFF As Long = 4
Private Const FVIRTKEY As Long = &H1
Private Const FSHIFT As Long = &H4
Private Const FALT As Long = &H10
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE As Long = &H0
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_PARENTNOTIFY As Long = &H210, WM_CREATE As Long = &H1
Private Const WM_STYLECHANGED As Long = &H7D
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
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_SETFONT As Long = &H30
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINT As Long = &H317, PRF_CLIENT As Long = &H4, PRF_ERASEBKGND As Long = &H8
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_DRAWITEM As Long = &H2B, ODT_TAB As Long = &H65
Private Const WM_USER As Long = &H400
Private Const TCS_SCROLLOPPOSITE As Long = &H1
Private Const TCS_BOTTOM As Long = &H2
Private Const TCS_RIGHT As Long = &H2
Private Const TCS_MULTISELECT As Long = &H4
Private Const TCS_FORCEICONLEFT As Long = &H10
Private Const TCS_FORCELABELLEFT As Long = &H20
Private Const TCS_HOTTRACK As Long = &H40
Private Const TCS_VERTICAL As Long = &H80
Private Const TCS_TABS As Long = &H0
Private Const TCS_BUTTONS As Long = &H100
Private Const TCS_FLATBUTTONS As Long = &H8
Private Const TCS_SINGLELINE As Long = &H0
Private Const TCS_MULTILINE As Long = &H200
Private Const TCS_RIGHTJUSTIFY As Long = &H0
Private Const TCS_FIXEDWIDTH As Long = &H400
Private Const TCS_RAGGEDRIGHT As Long = &H800
Private Const TCS_FOCUSONBUTTONDOWN As Long = &H1000
Private Const TCS_OWNERDRAWFIXED As Long = &H2000
Private Const TCS_TOOLTIPS As Long = &H4000
Private Const TCS_FOCUSNEVER As Long = &H8000&
Private Const TCS_EX_FLATSEPARATORS As Long = &H1
Private Const TCIF_TEXT As Long = &H1
Private Const TCIF_IMAGE As Long = &H2
Private Const TCIF_RTLREADING As Long = &H4
Private Const TCIF_PARAM As Long = &H8
Private Const TCIF_STATE As Long = &H10
Private Const TCIS_BUTTONPRESSED As Long = &H1
Private Const TCIS_HIGHLIGHTED As Long = &H2
Private Const TCM_FIRST As Long = &H1300
Private Const TCM_GETIMAGELIST As Long = (TCM_FIRST + 2)
Private Const TCM_SETIMAGELIST As Long = (TCM_FIRST + 3)
Private Const TCM_GETITEMCOUNT As Long = (TCM_FIRST + 4)
Private Const TCM_GETITEMA As Long = (TCM_FIRST + 5)
Private Const TCM_GETITEMW As Long = (TCM_FIRST + 60)
Private Const TCM_GETITEM As Long = TCM_GETITEMW
Private Const TCM_SETITEMA As Long = (TCM_FIRST + 6)
Private Const TCM_SETITEMW As Long = (TCM_FIRST + 61)
Private Const TCM_SETITEM As Long = TCM_SETITEMW
Private Const TCM_INSERTITEMA As Long = (TCM_FIRST + 7)
Private Const TCM_INSERTITEMW As Long = (TCM_FIRST + 62)
Private Const TCM_INSERTITEM As Long = TCM_INSERTITEMW
Private Const TCM_DELETEITEM As Long = (TCM_FIRST + 8)
Private Const TCM_DELETEALLITEMS As Long = (TCM_FIRST + 9)
Private Const TCM_GETITEMRECT As Long = (TCM_FIRST + 10)
Private Const TCM_GETCURSEL As Long = (TCM_FIRST + 11)
Private Const TCM_SETCURSEL As Long = (TCM_FIRST + 12)
Private Const TCM_HITTEST As Long = (TCM_FIRST + 13)
Private Const TCM_ADJUSTRECT As Long = (TCM_FIRST + 40)
Private Const TCM_SETITEMSIZE As Long = (TCM_FIRST + 41)
Private Const TCM_GETROWCOUNT As Long = (TCM_FIRST + 44)
Private Const TCM_GETTOOLTIPS As Long = (TCM_FIRST + 45)
Private Const TCM_SETTOOLTIPS As Long = (TCM_FIRST + 46)
Private Const TCM_GETCURFOCUS As Long = (TCM_FIRST + 47)
Private Const TCM_SETCURFOCUS As Long = (TCM_FIRST + 48)
Private Const TCM_SETMINTABWIDTH As Long = (TCM_FIRST + 49)
Private Const TCM_DESELECTALL As Long = (TCM_FIRST + 50)
Private Const TCM_HIGHLIGHTITEM As Long = (TCM_FIRST + 51)
Private Const TCM_SETEXTENDEDSTYLE As Long = (TCM_FIRST + 52)
Private Const TCM_GETEXTENDEDSTYLE As Long = (TCM_FIRST + 53)
Private Const TCHT_NOWHERE As Long = &H1
Private Const TCHT_ONITEMICON As Long = &H2
Private Const TCHT_ONITEMLABEL As Long = &H4
Private Const TCHT_ONITEM As Long = (TCHT_ONITEMICON Or TCHT_ONITEMLABEL)
Private Const MAX_PATH As Long = 260
Private Const TCN_FIRST As Long = (-550)
Private Const TCN_SELCHANGE As Long = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING As Long = (TCN_FIRST - 2)
Private Const TCN_FOCUSCHANGE As Long = (TCN_FIRST - 4)
Private Const TTF_RTLREADING As Long = &H4
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IOleControlVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private Type InitTabStruct
Caption As String
Key As String
Tag As String
ToolTipText As String
Image As Variant
ImageIndex As Long
End Type
Private TabStripHandle As LongPtr, TabStripToolTipHandle As LongPtr
Private TabStripAcceleratorHandle As LongPtr
Private TabStripFontHandle As LongPtr
Private TabStripBackColorBrush As LongPtr
Private TabStripTransparentBrush As LongPtr
Private TabStripCharCodeCache As Long
Private TabStripMouseOver As Boolean
Private TabStripDesignMode As Boolean
Private TabStripDoubleBufferEraseBkgDC As LongPtr
Private TabStripImageListObjectPointer As LongPtr, TabStripImageListHandle As LongPtr
Private TabStripStyleCache As Long
Private UCNoSetFocusFwd As Boolean
Private DispIdImageList As Long, ImageListArray() As String

#If ImplementPreTranslateMsg = True Then

Private Const UM_PRETRANSLATEMSG As Long = (WM_USER + 1100)
Private UsePreTranslateMsg As Boolean

#End If

Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropTabs As TbsTabs
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBackColor As OLE_COLOR
Private PropImageListName As String, PropImageListInit As Boolean
Private PropPlacement As TbsPlacementConstants
Private PropMultiRow As Boolean
Private PropMultiSelect As Boolean
Private PropHotTracking As Boolean
Private PropStyle As TbsStyleConstants
Private PropTabStyle As TbsTabStyleConstants
Private PropTabWidthStyle As TbsTabWidthStyleConstants
Private PropTabFixedWidth As Integer, PropTabFixedHeight As Integer
Private PropTabMinWidth As Integer
Private PropTabAlignment As TbsTabAlignmentConstants
Private PropSeparators As Boolean
Private PropShowTips As Boolean
Private PropDrawMode As TbsDrawModeConstants
Private PropTabScrollWheel As Boolean
Private PropDoubleBuffer As Boolean
Private PropTransparent As Boolean

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
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab, vbKeyReturn, vbKeyEscape
            If IsInputKey = True Then
                SendMessage hWnd, wMsg, wParam, ByVal lParam
                Handled = True
            End If
    End Select
End If
End Sub

#If VBA7 Then
Private Sub IOleControlVB_GetControlInfo(ByRef Handled As Boolean, ByRef AccelCount As Integer, ByRef AccelTable As LongPtr, ByRef Flags As Long)
#Else
Private Sub IOleControlVB_GetControlInfo(ByRef Handled As Boolean, ByRef AccelCount As Integer, ByRef AccelTable As Long, ByRef Flags As Long)
#End If
If TabStripAcceleratorHandle <> NULL_PTR Then
    DestroyAcceleratorTable TabStripAcceleratorHandle
    TabStripAcceleratorHandle = NULL_PTR
End If
If TabStripHandle <> NULL_PTR Then
    Dim Count As Long
    Count = CLng(SendMessage(TabStripHandle, TCM_GETITEMCOUNT, 0, ByVal 0&))
    If Count > 0 Then
        Dim i As Long, Accel As Integer, AccelArray() As TACCEL, AccelRefCount As Long
        For i = 1 To Count
            Accel = AccelCharCode(Me.FTabCaption(i))
            If Accel <> 0 Then
                ReDim Preserve AccelArray(0 To AccelRefCount) As TACCEL
                With AccelArray(AccelRefCount)
                .FVirt = FVIRTKEY Or FALT
                .Cmd = i
                .Key = (VkKeyScan(Accel) And &HFF&)
                End With
                AccelRefCount = AccelRefCount + 1
                ReDim Preserve AccelArray(0 To AccelRefCount) As TACCEL
                With AccelArray(AccelRefCount)
                .FVirt = FVIRTKEY Or FALT Or FSHIFT
                .Cmd = AccelArray(AccelRefCount - 1).Cmd
                .Key = AccelArray(AccelRefCount - 1).Key
                End With
                AccelRefCount = AccelRefCount + 1
            End If
        Next i
        If AccelRefCount > 0 Then
            AccelCount = AccelRefCount
            TabStripAcceleratorHandle = CreateAcceleratorTable(VarPtr(AccelArray(0)), AccelCount)
            AccelTable = TabStripAcceleratorHandle
            Flags = 0
            Handled = True
        End If
    End If
End If
End Sub

#If VBA7 Then
Private Sub IOleControlVB_OnMnemonic(ByRef Handled As Boolean, ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal Shift As Long)
#Else
Private Sub IOleControlVB_OnMnemonic(ByRef Handled As Boolean, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
#End If
If TabStripHandle <> NULL_PTR And wMsg = WM_SYSKEYDOWN Then
    Dim Accel As Long, Count As Long, i As Long
    Count = CLng(SendMessage(TabStripHandle, TCM_GETITEMCOUNT, 0, ByVal 0&))
    If Count > 0 Then
        For i = 1 To Count
            Accel = AccelCharCode(Me.FTabCaption(i))
            If (VkKeyScan(Accel) And &HFF&) = (wParam And &HFF&) Then
                If i <> SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&) - 1 Then Me.FTabSelected(i) = True
                Exit For
            End If
        Next i
        Handled = True
    End If
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispId As Long, ByRef DisplayName As String)
If DispId = DispIdImageList Then
    DisplayName = PropImageListName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispId As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispId = DispIdImageList Then
    On Error GoTo CATCH_EXCEPTION
    Call ComCtlsIPPBSetPredefinedStringsImageList(StringsOut(), CookiesOut(), UserControl.ParentControls, ImageListArray())
    On Error GoTo 0
    Handled = True
End If
Exit Sub
CATCH_EXCEPTION:
Handled = False
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispId As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispId = DispIdImageList Then
    If Cookie < UBound(ImageListArray()) Then Value = ImageListArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_TAB_CLASSES)

#If ImplementPreTranslateMsg = True Then

If SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject) = False Then UsePreTranslateMsg = True

#Else

Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)

#End If

Call SetVTableHandling(Me, VTableInterfaceControl)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIdImageList = 0 Then DispIdImageList = GetDispId(Me, "ImageList")
On Error Resume Next
TabStripDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBackColor = vbButtonFace
PropImageListName = "(None)"
PropPlacement = TbsPlacementTop
PropMultiRow = True
PropMultiSelect = False
PropHotTracking = True
PropStyle = TbsStyleTabs
PropTabStyle = TbsTabStyleStandard
PropTabWidthStyle = TbsTabWidthStyleJustified
PropTabFixedWidth = 0
PropTabFixedHeight = 0
PropTabMinWidth = (40 * PixelsPerDIP_X())
PropTabAlignment = TbsTabAlignmentStandard
PropSeparators = True
PropShowTips = False
PropDrawMode = TbsDrawModeNormal
PropTabScrollWheel = True
PropDoubleBuffer = True
PropTransparent = False
Call CreateTabStrip
Me.Tabs.Add
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIdImageList = 0 Then DispIdImageList = GetDispId(Me, "ImageList")
On Error Resume Next
TabStripDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
PropImageListName = .ReadProperty("ImageList", "(None)")
PropPlacement = .ReadProperty("Placement", TbsPlacementTop)
PropMultiRow = .ReadProperty("MultiRow", True)
PropMultiSelect = .ReadProperty("MultiSelect", False)
PropHotTracking = .ReadProperty("HotTracking", True)
PropStyle = .ReadProperty("Style", TbsTabStyleStandard)
PropTabStyle = .ReadProperty("TabStyle", TbsTabStyleStandard)
PropTabWidthStyle = .ReadProperty("TabWidthStyle", TbsTabWidthStyleJustified)
PropTabFixedWidth = (.ReadProperty("TabFixedWidth", 0) * PixelsPerDIP_X())
PropTabFixedHeight = (.ReadProperty("TabFixedHeight", 0) * PixelsPerDIP_Y())
PropTabMinWidth = (.ReadProperty("TabMinWidth", 40) * PixelsPerDIP_X())
PropTabAlignment = .ReadProperty("TabAlignment", TbsTabAlignmentStandard)
PropSeparators = .ReadProperty("Separators", True)
PropShowTips = .ReadProperty("ShowTips", False)
PropDrawMode = .ReadProperty("DrawMode", TbsDrawModeNormal)
PropTabScrollWheel = .ReadProperty("TabScrollWheel", True)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropTransparent = .ReadProperty("Transparent", False)
End With
With New PropertyBag
On Error Resume Next
.Contents = PropBag.ReadProperty("InitTabs", 0)
On Error GoTo 0
Dim InitTabsCount As Long, i As Long
Dim InitTabs() As InitTabStruct
InitTabsCount = .ReadProperty("InitTabsCount", 0)
If InitTabsCount > 0 Then
    ReDim InitTabs(1 To InitTabsCount) As InitTabStruct
    Dim VarValue As Variant
    For i = 1 To InitTabsCount
        InitTabs(i).Caption = .ReadProperty("InitTabsCaption" & CStr(i), vbNullString)
        InitTabs(i).Key = .ReadProperty("InitTabsKey" & CStr(i), vbNullString)
        InitTabs(i).Tag = .ReadProperty("InitTabsTag" & CStr(i), vbNullString)
        InitTabs(i).ToolTipText = .ReadProperty("InitTabsToolTipText" & CStr(i), vbNullString)
        VarValue = .ReadProperty("InitTabsImage" & CStr(i), 0)
        If VarType(VarValue) = vbArray + vbByte Then
            InitTabs(i).Image = VarToStr(VarValue)
            InitTabs(i).ImageIndex = .ReadProperty("InitTabsImageIndex" & CStr(i), 0)
        Else
            InitTabs(i).Image = VarValue
            InitTabs(i).ImageIndex = CLng(VarValue)
        End If
    Next i
End If
End With
Call CreateTabStrip
If InitTabsCount > 0 And TabStripHandle <> NULL_PTR Then
    Dim ImageListInit As Boolean
    ImageListInit = PropImageListInit
    PropImageListInit = True
    For i = 1 To InitTabsCount
        With Me.Tabs.Add(i, InitTabs(i).Key, InitTabs(i).Caption, InitTabs(i).ImageIndex)
        .FInit Me, InitTabs(i).Key, InitTabs(i).Image, InitTabs(i).ImageIndex
        .Tag = InitTabs(i).Tag
        .ToolTipText = InitTabs(i).ToolTipText
        End With
    Next i
    PropImageListInit = ImageListInit
End If
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "Placement", PropPlacement, TbsPlacementTop
.WriteProperty "MultiRow", PropMultiRow, True
.WriteProperty "MultiSelect", PropMultiSelect, False
.WriteProperty "HotTracking", PropHotTracking, True
.WriteProperty "Style", PropStyle, TbsTabStyleStandard
.WriteProperty "TabStyle", PropTabStyle, TbsTabStyleStandard
.WriteProperty "TabWidthStyle", PropTabWidthStyle, TbsTabWidthStyleJustified
.WriteProperty "TabFixedWidth", (PropTabFixedWidth / PixelsPerDIP_X()), 0
.WriteProperty "TabFixedHeight", (PropTabFixedHeight / PixelsPerDIP_Y()), 0
.WriteProperty "TabMinWidth", (PropTabMinWidth / PixelsPerDIP_X()), 40
.WriteProperty "TabAlignment", PropTabAlignment, TbsTabAlignmentStandard
.WriteProperty "Separators", PropSeparators, True
.WriteProperty "ShowTips", PropShowTips, False
.WriteProperty "DrawMode", PropDrawMode, TbsDrawModeNormal
.WriteProperty "TabScrollWheel", PropTabScrollWheel, True
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "Transparent", PropTransparent, False
End With
Dim Count As Long
Count = Me.Tabs.Count
With New PropertyBag
.WriteProperty "InitTabsCount", Count, 0
If Count > 0 Then
    Dim i As Long, VarValue As Variant
    For i = 1 To Count
        .WriteProperty "InitTabsCaption" & CStr(i), StrToVar(Me.Tabs(i).Caption), vbNullString
        .WriteProperty "InitTabsKey" & CStr(i), StrToVar(Me.Tabs(i).Key), vbNullString
        .WriteProperty "InitTabsTag" & CStr(i), StrToVar(Me.Tabs(i).Tag), vbNullString
        .WriteProperty "InitTabsToolTipText" & CStr(i), StrToVar(Me.Tabs(i).ToolTipText), vbNullString
        VarValue = Me.Tabs(i).Image
        If VarType(VarValue) = vbString Then
            .WriteProperty "InitTabsImage" & CStr(i), StrToVar(VarValue), 0
            .WriteProperty "InitTabsImageIndex" & CStr(i), Me.Tabs(i).ImageIndex, 0
        Else
            .WriteProperty "InitTabsImage" & CStr(i), VarValue, 0
        End If
    Next i
End If
PropBag.WriteProperty "InitTabs", .Contents, 0
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim P As POINTAPI
P.X = X
P.Y = Y
If TabStripHandle <> NULL_PTR Then MapWindowPoints UserControl.hWnd, TabStripHandle, P, 1
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim P As POINTAPI
P.X = X
P.Y = Y
If TabStripHandle <> NULL_PTR Then MapWindowPoints UserControl.hWnd, TabStripHandle, P, 1
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
If TabStripHandle <> NULL_PTR Then
    If PropTransparent = True Then
        MoveWindow TabStripHandle, 0, 0, .ScaleWidth, .ScaleHeight, 0
        If TabStripTransparentBrush <> NULL_PTR Then
            DeleteObject TabStripTransparentBrush
            TabStripTransparentBrush = NULL_PTR
        End If
        RedrawWindow TabStripHandle, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
    Else
        MoveWindow TabStripHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
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

Call RemoveVTableHandling(Me, VTableInterfaceControl)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyTabStrip
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    Me.ImageList = PropImageListName
    PropImageListInit = True
End If
TimerImageList.Enabled = False
End Sub

Public Property Get ControlsEnum() As VBRUN.ParentControls
Attribute ControlsEnum.VB_MemberFlags = "40"
Set ControlsEnum = UserControl.ParentControls
End Property

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
hWnd = TabStripHandle
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
Dim OldFontHandle As LongPtr
Set PropFont = NewFont
OldFontHandle = TabStripFontHandle
TabStripFontHandle = CreateGDIFontFromOLEFont(PropFont)
If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, WM_SETFONT, TabStripFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As LongPtr
OldFontHandle = TabStripFontHandle
TabStripFontHandle = CreateGDIFontFromOLEFont(PropFont)
If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, WM_SETFONT, TabStripFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If TabStripHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles TabStripHandle
    Else
        RemoveVisualStyles TabStripHandle
    End If
    Call SetVisualStylesUpDown
    Call SetVisualStylesToolTip
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If TabStripHandle <> NULL_PTR Then EnableWindow TabStripHandle, IIf(Value = True, 1, 0)
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
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
If TabStripDesignMode = False Then Call RefreshMousePointer
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
        If TabStripDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If TabStripDesignMode = False Then Call RefreshMousePointer
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
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
Attribute RightToLeft.VB_UserMemId = -611
RightToLeft = PropRightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
PropRightToLeft = Value
UserControl.RightToLeft = PropRightToLeft
Call ComCtlsCheckRightToLeft(PropRightToLeft, UserControl.RightToLeft, PropRightToLeftMode)
Dim dwMask As Long
If TabStripDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftLayout() As Boolean
Attribute RightToLeftLayout.VB_Description = "Returns/sets a value indicating if right-to-left mirror placement is turned on."
RightToLeftLayout = PropRightToLeftLayout
End Property

Public Property Let RightToLeftLayout(ByVal Value As Boolean)
PropRightToLeftLayout = Value
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftLayout"
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

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
If TabStripDesignMode = True Then
    If TabStripHandle <> NULL_PTR Then
        If Value = vbButtonFace And PropBackColor <> vbButtonFace Then
            Call ComCtlsRemoveSubclass(TabStripHandle)
            Call ComCtlsRemoveSubclass(UserControl.hWnd)
        ElseIf Value <> vbButtonFace And PropBackColor = vbButtonFace Then
            Call ComCtlsSetSubclass(TabStripHandle, Me, 3)
            Call ComCtlsSetSubclass(UserControl.hWnd, Me, 4)
        End If
    End If
End If
PropBackColor = Value
If TabStripHandle <> NULL_PTR Then
    If TabStripBackColorBrush <> NULL_PTR Then DeleteObject TabStripBackColorBrush
    If TabStripDesignMode = False Or PropBackColor <> vbButtonFace Then TabStripBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
End If
UserControl.BackColor = PropBackColor
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ImageList() As Variant
Attribute ImageList.VB_Description = "Returns/sets the image list control to be used."
If TabStripDesignMode = False Then
    If TabStripImageListHandle = NULL_PTR Then
        If PropImageListInit = False And TabStripImageListObjectPointer = NULL_PTR Then
            If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
            PropImageListInit = True
        End If
        Set ImageList = PropImageListControl
    Else
        ImageList = TabStripImageListHandle
    End If
Else
    ImageList = PropImageListName
End If
End Property

Public Property Set ImageList(ByVal Value As Variant)
Me.ImageList = Value
End Property

Public Property Let ImageList(ByVal Value As Variant)
If TabStripHandle <> NULL_PTR Then
    Dim Success As Boolean, Handle As LongPtr
    Select Case VarType(Value)
        Case vbObject
            If Not Value Is Nothing Then
                If TypeName(Value) = "ImageList" Then
                    On Error Resume Next
                    Handle = Value.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
                    On Error GoTo 0
                Else
                    Err.Raise Number:=35610, Description:="Invalid object"
                End If
            End If
            If Success = True Then
                SendMessage TabStripHandle, TCM_SETIMAGELIST, 0, ByVal Handle
                TabStripImageListObjectPointer = ObjPtr(Value)
                TabStripImageListHandle = NULL_PTR
                PropImageListName = ProperControlName(Value)
            End If
        Case vbString
            On Error Resume Next
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In UserControl.ParentControls
                If TypeName(ControlEnum) = "ImageList" Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = Value And Not CompareName = vbNullString Then
                        Err.Clear
                        Handle = ControlEnum.hImageList
                        Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
                        If Success = True Then
                            SendMessage TabStripHandle, TCM_SETIMAGELIST, 0, ByVal Handle
                            If TabStripDesignMode = False Then
                                TabStripImageListObjectPointer = ObjPtr(ControlEnum)
                                TabStripImageListHandle = NULL_PTR
                            End If
                            PropImageListName = Value
                            Exit For
                        ElseIf TabStripDesignMode = True Then
                            PropImageListName = Value
                            Success = True
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
            On Error GoTo 0
        Case vbLong, &H14 ' vbLongLong
            Handle = Value
            Success = CBool(Handle <> NULL_PTR)
            If Success = True Then
                SendMessage TabStripHandle, TCM_SETIMAGELIST, 0, ByVal Handle
                TabStripImageListObjectPointer = NULL_PTR
                TabStripImageListHandle = Handle
                PropImageListName = "(None)"
            End If
        Case Else
            Err.Raise 13
    End Select
    If Success = False Then
        SendMessage TabStripHandle, TCM_SETIMAGELIST, 0, ByVal 0&
        TabStripImageListObjectPointer = NULL_PTR
        TabStripImageListHandle = NULL_PTR
        PropImageListName = "(None)"
    ElseIf Handle = NULL_PTR Then
        SendMessage TabStripHandle, TCM_SETIMAGELIST, 0, ByVal 0&
    End If
End If
If PropMultiRow = False Then Call SetVisualStylesUpDown
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get Placement() As TbsPlacementConstants
Attribute Placement.VB_Description = "Returns/sets a value that indicates on which side of the control the tabs will be displayed. This property is ignored if the version of comctl32.dll is 6.0 or higher."
Placement = PropPlacement
End Property

Public Property Let Placement(ByVal Value As TbsPlacementConstants)
Select Case Value
    Case TbsPlacementTop, TbsPlacementBottom, TbsPlacementLeft, TbsPlacementRight
        PropPlacement = Value
    Case Else
        Err.Raise 380
End Select
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "Placement"
End Property

Public Property Get MultiRow() As Boolean
Attribute MultiRow.VB_Description = "Returns/sets a value indicating whether the control can display more than one row of tabs. This flag is always set to true when the tab style property is set to opposite."
MultiRow = PropMultiRow
End Property

Public Property Let MultiRow(ByVal Value As Boolean)
PropMultiRow = Value
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "MultiRow"
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether or not multiple tabs can be selected by holding down the CTRL key when clicking. Only applicable if the style property is set to buttons or flat buttons."
MultiSelect = PropMultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As Boolean)
PropMultiSelect = Value
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "MultiSelect"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value that determines whether or not the control highlights the tabs as the pointer passes over them. The flag is ignored on Windows XP (or above) when the desktop theme overrides it."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get Style() As TbsStyleConstants
Attribute Style.VB_Description = "Returns/sets the style appearance."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As TbsStyleConstants)
Select Case Value
    Case TbsStyleTabs, TbsStyleButtons, TbsStyleFlatButtons
        PropStyle = Value
    Case Else
        Err.Raise 380
End Select
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "Style"
End Property

Public Property Get TabStyle() As TbsTabStyleConstants
Attribute TabStyle.VB_Description = "Returns/sets a value that determines how remaining rows of tabs in front of a selected tab are repositioned."
TabStyle = PropTabStyle
End Property

Public Property Let TabStyle(ByVal Value As TbsTabStyleConstants)
Select Case Value
    Case TbsTabStyleStandard, TbsTabStyleOpposite
        PropTabStyle = Value
    Case Else
        Err.Raise 380
End Select
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "TabStyle"
End Property

Public Property Get TabWidthStyle() As TbsTabWidthStyleConstants
Attribute TabWidthStyle.VB_Description = "Returns/sets the width and justification of all tabs."
TabWidthStyle = PropTabWidthStyle
End Property

Public Property Let TabWidthStyle(ByVal Value As TbsTabWidthStyleConstants)
Select Case Value
    Case TbsTabWidthStyleJustified, TbsTabWidthStyleNonJustified, TbsTabWidthStyleFixed
        PropTabWidthStyle = Value
    Case Else
        Err.Raise 380
End Select
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "TabWidthStyle"
End Property

Public Property Get TabFixedWidth() As Single
Attribute TabFixedWidth.VB_Description = "Returns/sets a fixed width of a tab, but only if the tab width style property is set to fixed."
TabFixedWidth = UserControl.ScaleX(PropTabFixedWidth, vbPixels, vbContainerSize)
End Property

Public Property Let TabFixedWidth(ByVal Value As Single)
If Value < 0 Then
    If TabStripDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer, ErrValue As Long
On Error Resume Next
IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
ErrValue = Err.Number
On Error GoTo 0
If IntValue >= 0 And ErrValue = 0 Then
    PropTabFixedWidth = IntValue
    If PropTabWidthStyle = TbsTabWidthStyleFixed Then
        If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, TCM_SETITEMSIZE, 0, ByVal MakeDWord(PropTabFixedWidth, PropTabFixedHeight)
    End If
Else
    If TabStripDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "TabFixedWidth"
End Property

Public Property Get TabFixedHeight() As Single
Attribute TabFixedHeight.VB_Description = "Returns/sets a fixed height of a tab, but only if the tab width style property is set to fixed."
TabFixedHeight = UserControl.ScaleY(PropTabFixedHeight, vbPixels, vbContainerSize)
End Property

Public Property Let TabFixedHeight(ByVal Value As Single)
If Value < 0 Then
    If TabStripDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer, ErrValue As Long
On Error Resume Next
IntValue = CInt(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
ErrValue = Err.Number
On Error GoTo 0
If IntValue >= 0 And ErrValue = 0 Then
    PropTabFixedHeight = IntValue
    If PropTabWidthStyle = TbsTabWidthStyleFixed Then
        If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, TCM_SETITEMSIZE, 0, ByVal MakeDWord(PropTabFixedWidth, PropTabFixedHeight)
    End If
Else
    If TabStripDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "TabFixedHeight"
End Property

Public Property Get TabMinWidth() As Single
Attribute TabMinWidth.VB_Description = "Returns/sets a minimum width of a tab."
If PropTabMinWidth <> -1 Then
    TabMinWidth = UserControl.ScaleX(PropTabMinWidth, vbPixels, vbContainerSize)
Else
    TabMinWidth = -1
End If
End Property

Public Property Let TabMinWidth(ByVal Value As Single)
If Value < 0 And Not Value = -1 Then
    If TabStripDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer, ErrValue As Long
On Error Resume Next
If Value <> -1 Then IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels)) Else IntValue = -1
ErrValue = Err.Number
On Error GoTo 0
If (IntValue >= 0 Or IntValue = -1) And ErrValue = 0 Then
    PropTabMinWidth = IntValue
    If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, TCM_SETMINTABWIDTH, 0, ByVal CLng(PropTabMinWidth)
Else
    If TabStripDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "TabMinWidth"
End Property

Public Property Get TabAlignment() As TbsTabAlignmentConstants
Attribute TabAlignment.VB_Description = "Returns/sets the tab alignment, but only if the tab width style property is set to fixed."
TabAlignment = PropTabAlignment
End Property

Public Property Let TabAlignment(ByVal Value As TbsTabAlignmentConstants)
Select Case Value
    Case TbsTabAlignmentStandard, TbsTabAlignmentImageLeft, TbsTabAlignmentImageCaptionLeft
        PropTabAlignment = Value
    Case Else
        Err.Raise 380
End Select
If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
UserControl.PropertyChanged "TabAlignment"
End Property

Public Property Get Separators() As Boolean
Attribute Separators.VB_Description = "Returns/sets a value that determines whether or not the control will draw separators between the tabs. Only applicable if the style property is set to flat buttons."
If TabStripHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(TabStripHandle, TCM_GETEXTENDEDSTYLE, 0, ByVal 0&))
    Separators = CBool((dwStyle And TCS_EX_FLATSEPARATORS) = TCS_EX_FLATSEPARATORS)
Else
    Separators = PropSeparators
End If
End Property

Public Property Let Separators(ByVal Value As Boolean)
PropSeparators = Value
If TabStripHandle <> NULL_PTR Then
    If PropSeparators = False Then
        SendMessage TabStripHandle, TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, ByVal 0&
    Else
        SendMessage TabStripHandle, TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, ByVal TCS_EX_FLATSEPARATORS
    End If
End If
UserControl.PropertyChanged "Separators"
End Property

Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowTips = PropShowTips
End Property

Public Property Let ShowTips(ByVal Value As Boolean)
PropShowTips = Value
If TabStripHandle <> NULL_PTR And TabStripDesignMode = False Then
    If PropShowTips = False Then
        SendMessage TabStripHandle, TCM_SETTOOLTIPS, 0, ByVal 0&
    Else
        If TabStripToolTipHandle <> NULL_PTR Then
            SendMessage TabStripHandle, TCM_SETTOOLTIPS, TabStripToolTipHandle, ByVal 0&
        Else
            Call ReCreateTabStrip
        End If
    End If
End If
UserControl.PropertyChanged "ShowTips"
End Property

Public Property Get DrawMode() As TbsDrawModeConstants
Attribute DrawMode.VB_Description = "Returns/sets a value indicating whether your code or the operating system will handle drawing of the elements."
DrawMode = PropDrawMode
End Property

Public Property Let DrawMode(ByVal Value As TbsDrawModeConstants)
Select Case Value
    Case TbsDrawModeNormal, TbsDrawModeOwnerDrawFixed
        If TabStripDesignMode = False Then
            Err.Raise Number:=382, Description:="DrawMode property is read-only at run time"
        Else
            PropDrawMode = Value
            If TabStripHandle <> NULL_PTR Then Call ReCreateTabStrip
        End If
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "DrawMode"
End Property

Public Property Get TabScrollWheel() As Boolean
Attribute TabScrollWheel.VB_Description = "Returns/sets a value that determines whether or not the selected tab can be switched using the mouse scroll wheel."
TabScrollWheel = PropTabScrollWheel
End Property

Public Property Let TabScrollWheel(ByVal Value As Boolean)
PropTabScrollWheel = Value
UserControl.PropertyChanged "TabScrollWheel"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value indicating if the background is a replica of the underlying background to simulate transparency."
Transparent = PropTransparent
End Property

Public Property Let Transparent(ByVal Value As Boolean)
PropTransparent = Value
Me.Refresh
UserControl.PropertyChanged "Transparent"
End Property

Public Property Get Tabs() As TbsTabs
Attribute Tabs.VB_Description = "Returns a reference to a collection of the tab objects."
If PropTabs Is Nothing Then
    Set PropTabs = New TbsTabs
    PropTabs.FInit Me
End If
Set Tabs = PropTabs
End Property

Friend Sub FTabsAdd(ByVal Index As Long, Optional ByVal Caption As String, Optional ByVal ImageIndex As Long)
Dim TabIndex As Long
Dim TCI As TCITEM
With TCI
.Mask = TCIF_TEXT Or TCIF_IMAGE Or TCIF_PARAM
If PropRightToLeft = True And PropRightToLeftLayout = False Then .Mask = .Mask Or TCIF_RTLREADING
.iImage = ImageIndex - 1
.cchTextMax = Len(Caption)
.pszText = StrPtr(Caption)
.lParam = 0
End With
If Index = 0 Then
    TabIndex = Me.Tabs.Count + 1
Else
    TabIndex = Index
End If
If TabStripHandle <> NULL_PTR Then
    SendMessage TabStripHandle, TCM_INSERTITEM, TabIndex - 1, ByVal VarPtr(TCI)
    Call OnControlInfoChanged(Me)
End If
If PropMultiRow = False Then Call SetVisualStylesUpDown
UserControl.PropertyChanged "InitTabs"
End Sub

Friend Sub FTabsRemove(ByVal Index As Long)
If TabStripHandle <> NULL_PTR Then
    SendMessage TabStripHandle, TCM_DELETEITEM, Index - 1, ByVal 0&
    Call OnControlInfoChanged(Me)
End If
UserControl.PropertyChanged "InitTabs"
End Sub

Friend Sub FTabsClear()
If TabStripHandle <> NULL_PTR Then
    SendMessage TabStripHandle, TCM_DELETEALLITEMS, 0, ByVal 0&
    Call OnControlInfoChanged(Me)
End If
End Sub

Friend Property Get FTabCaption(ByVal Index As Long) As String
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM, Buffer As String
    With TCI
    Buffer = String(MAX_PATH, vbNullChar) & vbNullChar
    .Mask = TCIF_TEXT
    .pszText = StrPtr(Buffer)
    .cchTextMax = Len(Buffer)
    SendMessage TabStripHandle, TCM_GETITEM, Index - 1, ByVal VarPtr(TCI)
    FTabCaption = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End With
End If
End Property

Friend Property Let FTabCaption(ByVal Index As Long, ByVal Value As String)
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM
    With TCI
    .Mask = TCIF_TEXT
    If PropRightToLeft = True And PropRightToLeftLayout = False Then .Mask = .Mask Or TCIF_RTLREADING
    .pszText = StrPtr(Value)
    .cchTextMax = Len(Value)
    SendMessage TabStripHandle, TCM_SETITEM, Index - 1, ByVal VarPtr(TCI)
    Call OnControlInfoChanged(Me)
    End With
End If
End Property

Friend Property Get FTabImage(ByVal Index As Long) As Long
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM
    With TCI
    .Mask = TCIF_IMAGE
    SendMessage TabStripHandle, TCM_GETITEM, Index - 1, ByVal VarPtr(TCI)
    FTabImage = .iImage + 1
    End With
End If
End Property

Friend Property Let FTabImage(ByVal Index As Long, ByVal Value As Long)
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM
    With TCI
    .Mask = TCIF_IMAGE
    .iImage = Value - 1
    SendMessage TabStripHandle, TCM_SETITEM, Index - 1, ByVal VarPtr(TCI)
    End With
End If
End Property

Friend Property Get FTabSelected(ByVal Index As Long) As Boolean
If TabStripHandle <> NULL_PTR Then
    Dim SelIndex As Long
    SelIndex = CLng(SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&))
    If SelIndex > -1 Then FTabSelected = CBool((SelIndex + 1) = Index)
End If
End Property

Friend Property Let FTabSelected(ByVal Index As Long, ByVal Value As Boolean)
If TabStripHandle <> NULL_PTR Then
    If Value = True Then
        Dim Cancel As Boolean
        RaiseEvent TabBeforeClick(Me.Tabs(Index), Cancel)
        If Cancel = False Then
            SendMessage TabStripHandle, TCM_SETCURSEL, Index - 1, ByVal 0&
            RaiseEvent TabClick(Me.Tabs(Index))
        End If
    Else
        If SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&) = Index - 1 Then SendMessage TabStripHandle, TCM_SETCURSEL, -1, ByVal 0&
    End If
End If
End Property

Friend Property Get FTabPressed(ByVal Index As Long) As Boolean
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM
    With TCI
    .Mask = TCIF_STATE
    .dwStateMask = TCIS_BUTTONPRESSED
    SendMessage TabStripHandle, TCM_GETITEM, Index - 1, ByVal VarPtr(TCI)
    FTabPressed = CBool((.dwState And TCIS_BUTTONPRESSED) = TCIS_BUTTONPRESSED)
    End With
End If
End Property

Friend Property Let FTabPressed(ByVal Index As Long, ByVal Value As Boolean)
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM
    With TCI
    .Mask = TCIF_STATE
    .dwStateMask = TCIS_BUTTONPRESSED
    If Value = True Then
        .dwState = TCIS_BUTTONPRESSED
    Else
        .dwState = 0
    End If
    SendMessage TabStripHandle, TCM_SETITEM, Index - 1, ByVal VarPtr(TCI)
    End With
End If
End Property

Friend Property Get FTabHighLighted(ByVal Index As Long) As Boolean
If TabStripHandle <> NULL_PTR Then
    Dim TCI As TCITEM
    With TCI
    .Mask = TCIF_STATE
    .dwStateMask = TCIS_HIGHLIGHTED
    SendMessage TabStripHandle, TCM_GETITEM, Index - 1, ByVal VarPtr(TCI)
    FTabHighLighted = CBool((.dwState And TCIS_HIGHLIGHTED) = TCIS_HIGHLIGHTED)
    End With
End If
End Property

Friend Property Let FTabHighLighted(ByVal Index As Long, ByVal Value As Boolean)
If TabStripHandle <> NULL_PTR Then
    If Value = True Then
        SendMessage TabStripHandle, TCM_HIGHLIGHTITEM, Index - 1, ByVal 1&
    Else
        SendMessage TabStripHandle, TCM_HIGHLIGHTITEM, Index - 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FTabLeft(ByVal Index As Long) As Single
If TabStripHandle <> NULL_PTR Then
    Dim RC As RECT
    SendMessage TabStripHandle, TCM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FTabLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FTabTop(ByVal Index As Long) As Single
If TabStripHandle <> NULL_PTR Then
    Dim RC As RECT
    SendMessage TabStripHandle, TCM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FTabTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FTabWidth(ByVal Index As Long) As Single
If TabStripHandle <> NULL_PTR Then
    Dim RC As RECT
    SendMessage TabStripHandle, TCM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FTabWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FTabHeight(ByVal Index As Long) As Single
If TabStripHandle <> NULL_PTR Then
    Dim RC As RECT
    SendMessage TabStripHandle, TCM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FTabHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Private Sub CreateTabStrip()
If TabStripHandle <> NULL_PTR Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or TCS_FOCUSONBUTTONDOWN
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropTabStyle = TbsTabStyleOpposite Then
    PropStyle = TbsStyleTabs
    PropMultiRow = True
End If
If ComCtlsSupportLevel() = 0 Then
    Select Case PropPlacement
        Case TbsPlacementBottom
            dwStyle = dwStyle Or TCS_BOTTOM
        Case TbsPlacementLeft, TbsPlacementRight
            PropMultiRow = True
            dwStyle = dwStyle Or TCS_VERTICAL
            If PropPlacement = TbsPlacementRight Then dwStyle = dwStyle Or TCS_RIGHT
    End Select
End If
If PropMultiRow = True Then
    dwStyle = dwStyle Or TCS_MULTILINE
Else
    dwStyle = dwStyle Or TCS_SINGLELINE
End If
If PropMultiSelect = True Then dwStyle = dwStyle Or TCS_MULTISELECT
If PropHotTracking = True Then dwStyle = dwStyle Or TCS_HOTTRACK
Select Case PropStyle
    Case TbsStyleTabs
        dwStyle = dwStyle Or TCS_TABS
    Case TbsStyleButtons
        dwStyle = dwStyle Or TCS_BUTTONS
    Case TbsStyleFlatButtons
        dwStyle = dwStyle Or TCS_BUTTONS Or TCS_FLATBUTTONS
End Select
If PropTabStyle = TbsTabStyleOpposite Then dwStyle = dwStyle Or TCS_SCROLLOPPOSITE
Select Case PropTabWidthStyle
    Case TbsTabWidthStyleJustified
        PropTabAlignment = TbsTabAlignmentStandard
        dwStyle = dwStyle Or TCS_RIGHTJUSTIFY
    Case TbsTabWidthStyleNonJustified
        PropTabAlignment = TbsTabAlignmentStandard
        dwStyle = dwStyle Or TCS_RAGGEDRIGHT
    Case TbsTabWidthStyleFixed
        dwStyle = dwStyle Or TCS_FIXEDWIDTH
End Select
If PropTabWidthStyle = TbsTabWidthStyleFixed Then
    Select Case PropTabAlignment
        Case TbsTabAlignmentImageLeft
            dwStyle = dwStyle Or TCS_FORCEICONLEFT
        Case TbsTabAlignmentImageCaptionLeft
            ' TCS_FORCELABELLEFT implies the TCS_FORCEICONLEFT style.
            dwStyle = dwStyle Or TCS_FORCELABELLEFT
    End Select
End If
If PropShowTips = True And TabStripDesignMode = False Then dwStyle = dwStyle Or TCS_TOOLTIPS
If PropDrawMode = TbsDrawModeOwnerDrawFixed Then dwStyle = dwStyle Or TCS_OWNERDRAWFIXED
If TabStripDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
TabStripHandle = CreateWindowEx(dwExStyle, StrPtr("SysTabControl32"), NULL_PTR, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If TabStripHandle <> NULL_PTR Then
    Call ComCtlsShowAllUIStates(TabStripHandle)
    TabStripToolTipHandle = SendMessage(TabStripHandle, TCM_GETTOOLTIPS, 0, ByVal 0&)
    If TabStripToolTipHandle <> NULL_PTR Then Call ComCtlsInitToolTip(TabStripToolTipHandle)
    If PropTabWidthStyle = TbsTabWidthStyleFixed Then SendMessage TabStripHandle, TCM_SETITEMSIZE, 0, ByVal MakeDWord(PropTabFixedWidth, PropTabFixedHeight)
    SendMessage TabStripHandle, TCM_SETMINTABWIDTH, 0, ByVal CLng(PropTabMinWidth)
    Me.Refresh
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Separators = PropSeparators
If TabStripDesignMode = False Then
    If TabStripHandle <> NULL_PTR Then
        If TabStripBackColorBrush = NULL_PTR Then TabStripBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(TabStripHandle, Me, 1)
        TabStripStyleCache = dwStyle
    End If
    
    #If ImplementPreTranslateMsg = True Then
    
    If UsePreTranslateMsg = True Then Call ComCtlsPreTranslateMsgAddHook
    
    #End If
    
ElseIf PropBackColor <> vbButtonFace Then
    If TabStripHandle <> NULL_PTR Then
        If TabStripBackColorBrush = NULL_PTR Then TabStripBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(TabStripHandle, Me, 3)
        Call ComCtlsSetSubclass(UserControl.hWnd, Me, 4)
    End If
End If
UserControl.BackColor = PropBackColor
End Sub

Private Sub ReCreateTabStrip()
Dim Locked As Boolean
With Me
Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
Dim ReInitTabsCount As Long
Dim ReInitTabs() As InitTabStruct
ReInitTabsCount = .Tabs.Count
If ReInitTabsCount > 0 Then
    ReDim ReInitTabs(1 To ReInitTabsCount) As InitTabStruct
    Dim i As Long
    For i = 1 To ReInitTabsCount
        With .Tabs(i)
        ReInitTabs(i).Caption = .Caption
        ReInitTabs(i).Key = .Key
        ReInitTabs(i).Tag = .Tag
        ReInitTabs(i).ToolTipText = .ToolTipText
        ReInitTabs(i).Image = .Image
        ReInitTabs(i).ImageIndex = .ImageIndex
        End With
    Next i
End If
Dim CurrIndex As Long
If TabStripHandle <> NULL_PTR Then CurrIndex = CLng(SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&)) + 1
.Tabs.Clear
Call DestroyTabStrip
Call CreateTabStrip
Call UserControl_Resize
If TabStripDesignMode = False Then
    If Not PropImageListControl Is Nothing Then Set .ImageList = PropImageListControl
Else
    If Not PropImageListName = "(None)" Then .ImageList = PropImageListName
End If
If ReInitTabsCount > 0 Then
    For i = 1 To ReInitTabsCount
        With .Tabs.Add(i, ReInitTabs(i).Key, ReInitTabs(i).Caption, ReInitTabs(i).ImageIndex)
        .FInit Me, ReInitTabs(i).Key, ReInitTabs(i).Image, ReInitTabs(i).ImageIndex
        .Tag = ReInitTabs(i).Tag
        .ToolTipText = ReInitTabs(i).ToolTipText
        End With
    Next i
End If
If PropMultiRow = False Then Call SetVisualStylesUpDown
If TabStripHandle <> NULL_PTR Then If CurrIndex <> 0 Then SendMessage TabStripHandle, TCM_SETCURSEL, CurrIndex - 1, ByVal 0&
If Locked = True Then LockWindowUpdate 0
.Refresh
End With
End Sub

Private Sub DestroyTabStrip()
If TabStripHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(TabStripHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
If TabStripDesignMode = False Then
    
    #If ImplementPreTranslateMsg = True Then
    
    If UsePreTranslateMsg = True Then Call ComCtlsPreTranslateMsgReleaseHook
    
    #End If
    
End If
ShowWindow TabStripHandle, SW_HIDE
SetParent TabStripHandle, NULL_PTR
DestroyWindow TabStripHandle
TabStripHandle = NULL_PTR
TabStripToolTipHandle = NULL_PTR
If TabStripFontHandle <> NULL_PTR Then
    DeleteObject TabStripFontHandle
    TabStripFontHandle = NULL_PTR
End If
If TabStripAcceleratorHandle <> NULL_PTR Then
    DestroyAcceleratorTable TabStripAcceleratorHandle
    TabStripAcceleratorHandle = NULL_PTR
End If
If TabStripBackColorBrush <> NULL_PTR Then
    DeleteObject TabStripBackColorBrush
    TabStripBackColorBrush = NULL_PTR
End If
If TabStripTransparentBrush <> NULL_PTR Then
    DeleteObject TabStripTransparentBrush
    TabStripTransparentBrush = NULL_PTR
End If
TabStripStyleCache = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If TabStripTransparentBrush <> NULL_PTR Then
    DeleteObject TabStripTransparentBrush
    TabStripTransparentBrush = NULL_PTR
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get ClientLeft() As Single
Attribute ClientLeft.VB_Description = "Returns the left coordinate of the internal area."
Attribute ClientLeft.VB_MemberFlags = "400"
Dim RC As RECT
If TabStripHandle <> NULL_PTR Then
    GetClientRect TabStripHandle, RC
    SendMessage TabStripHandle, TCM_ADJUSTRECT, 0, ByVal VarPtr(RC)
End If
ClientLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End Property

Public Property Get ClientTop() As Single
Attribute ClientTop.VB_Description = "Returns the top coordinate of the internal area."
Attribute ClientTop.VB_MemberFlags = "400"
Dim RC As RECT
If TabStripHandle <> NULL_PTR Then
    GetClientRect TabStripHandle, RC
    SendMessage TabStripHandle, TCM_ADJUSTRECT, 0, ByVal VarPtr(RC)
End If
ClientTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End Property

Public Property Get ClientWidth() As Single
Attribute ClientWidth.VB_Description = "Returns/sets the width of the internal area."
Attribute ClientWidth.VB_MemberFlags = "400"
Dim RC As RECT
If TabStripHandle <> NULL_PTR Then
    GetClientRect TabStripHandle, RC
    SendMessage TabStripHandle, TCM_ADJUSTRECT, 0, ByVal VarPtr(RC)
End If
ClientWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End Property

Public Property Let ClientWidth(ByVal Value As Single)
Dim RC As RECT
RC.Right = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, TCM_ADJUSTRECT, 1, ByVal VarPtr(RC)
With UserControl
.Extender.Move .Extender.Left, .Extender.Top, .ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize), .Extender.Height
End With
End Property

Public Property Get ClientHeight() As Single
Attribute ClientHeight.VB_Description = "Returns/sets the height of the internal area."
Attribute ClientHeight.VB_MemberFlags = "400"
Dim RC As RECT
If TabStripHandle <> NULL_PTR Then
    GetClientRect TabStripHandle, RC
    SendMessage TabStripHandle, TCM_ADJUSTRECT, 0, ByVal VarPtr(RC)
End If
ClientHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End Property

Public Property Let ClientHeight(ByVal Value As Single)
Dim RC As RECT
RC.Bottom = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, TCM_ADJUSTRECT, 1, ByVal VarPtr(RC)
With UserControl
.Extender.Move .Extender.Left, .Extender.Top, .Extender.Width, .ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End With
End Property

Public Sub DeselectAll()
Attribute DeselectAll.VB_Description = "Resets the selected state for all tabs. This is only meaningful if the style property is set to buttons or flat buttons."
If TabStripHandle <> NULL_PTR Then SendMessage TabStripHandle, TCM_DESELECTALL, 0, ByVal 0&
End Sub

Public Property Get RowCount() As Long
Attribute RowCount.VB_Description = "Retrieves the current number of rows of tabs. Only tab strip controls that have the multiline property set to true can have can have multiple rows of tabs."
If TabStripHandle <> NULL_PTR Then RowCount = CLng(SendMessage(TabStripHandle, TCM_GETROWCOUNT, 0, ByVal 0&))
End Property

Public Property Get SelectedItem() As TbsTab
Attribute SelectedItem.VB_Description = "Returns/sets the selected tab."
Attribute SelectedItem.VB_MemberFlags = "400"
If TabStripHandle <> NULL_PTR Then
    Dim SelIndex As Long
    SelIndex = CLng(SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&))
    If SelIndex > -1 Then Set SelectedItem = Me.Tabs(SelIndex + 1)
End If
End Property

Public Property Let SelectedItem(ByVal Value As TbsTab)
Set Me.SelectedItem = Value
End Property

Public Property Set SelectedItem(ByVal Value As TbsTab)
If TabStripHandle <> NULL_PTR Then
    If Not Value Is Nothing Then
        Value.Selected = True
    Else
        SendMessage TabStripHandle, TCM_SETCURSEL, -1, ByVal 0&
    End If
End If
End Property

#If VBA7 Then
Public Sub DrawBackground(ByVal hWnd As LongPtr, ByVal hDC As LongPtr)
Attribute DrawBackground.VB_Description = "Draws the background to a given device context (DC) to a specified window."
#Else
Public Sub DrawBackground(ByVal hWnd As Long, ByVal hDC As Long)
Attribute DrawBackground.VB_Description = "Draws the background to a given device context (DC) to a specified window."
#End If
If TabStripHandle <> NULL_PTR And hWnd <> 0 And hDC <> 0 Then
    Dim RC As RECT, P As POINTAPI
    GetClientRect hWnd, RC
    MapWindowPoints hWnd, TabStripHandle, RC, 2
    P.X = RC.Left
    P.Y = RC.Top
    SetViewportOrgEx hDC, -P.X, -P.Y, P
    SendMessage TabStripHandle, WM_PRINT, hDC, ByVal PRF_CLIENT Or PRF_ERASEBKGND
    SetViewportOrgEx hDC, P.X, P.Y, P
End If
End Sub

Public Function HitTest(ByVal X As Single, ByVal Y As Single, Optional ByRef HitResult As TbsHitResultConstants) As TbsTab
Attribute HitTest.VB_Description = "Returns a reference to the tab object located at the coordinates of X and Y."
If TabStripHandle <> NULL_PTR Then
    Dim TCHTI As TCHITTESTINFO, Index As Long
    With TCHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    Index = CLng(SendMessage(TabStripHandle, TCM_HITTEST, 0, ByVal VarPtr(TCHTI))) + 1
    If Index > 0 Then
        Set HitTest = Me.Tabs(Index)
        Select Case .Flags
            Case TCHT_NOWHERE
                HitResult = TbsHitResultNoWhere
            Case TCHT_ONITEM
                HitResult = TbsHitResultItem
            Case TCHT_ONITEMICON
                HitResult = TbsHitResultItemIcon
            Case TCHT_ONITEMLABEL
                HitResult = TbsHitResultItemLabel
        End Select
    End If
    End With
End If
End Function

Private Function CreateTransparentBrush(ByVal hDC As LongPtr) As LongPtr
Dim hDCBmp As LongPtr
Dim hBmp As LongPtr, hBmpOld As LongPtr
With UserControl
hDCBmp = CreateCompatibleDC(hDC)
If hDCBmp <> NULL_PTR Then
    hBmp = CreateCompatibleBitmap(hDC, .ScaleWidth, .ScaleHeight)
    If hBmp <> NULL_PTR Then
        hBmpOld = SelectObject(hDCBmp, hBmp)
        Dim WndRect As RECT, P As POINTAPI
        GetWindowRect .hWnd, WndRect
        MapWindowPoints HWND_DESKTOP, GetParent(.hWnd), WndRect, 2
        P.X = WndRect.Left
        P.Y = WndRect.Top
        SetViewportOrgEx hDCBmp, -P.X, -P.Y, P
        SendMessage GetParent(.hWnd), WM_PAINT, hDCBmp, ByVal 0&
        SetViewportOrgEx hDCBmp, P.X, P.Y, P
        CreateTransparentBrush = CreatePatternBrush(hBmp)
        SelectObject hDCBmp, hBmpOld
        DeleteObject hBmp
    End If
    DeleteDC hDCBmp
End If
End With
End Function

Private Sub SetVisualStylesUpDown()
If TabStripHandle <> NULL_PTR Then
    Dim UpDownHandle As LongPtr
    UpDownHandle = FindWindowEx(TabStripHandle, NULL_PTR, StrPtr("msctls_updown32"), NULL_PTR)
    If UpDownHandle <> NULL_PTR And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles UpDownHandle
        Else
            RemoveVisualStyles UpDownHandle
        End If
    End If
End If
End Sub

Private Sub SetVisualStylesToolTip()
If TabStripHandle <> NULL_PTR Then
    If TabStripToolTipHandle <> NULL_PTR And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles TabStripToolTipHandle
        Else
            RemoveVisualStyles TabStripToolTipHandle
        End If
    End If
End If
End Sub

Private Function PropImageListControl() As Object
If TabStripImageListObjectPointer <> NULL_PTR Then Set PropImageListControl = PtrToObj(TabStripImageListObjectPointer)
End Function

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
    Case 3
        ISubclass_Message = WindowProcControlDesignMode(hWnd, wMsg, wParam, lParam)
    Case 4
        ISubclass_Message = WindowProcUserControlDesignMode(hWnd, wMsg, wParam, lParam)
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
        
    Case WM_LBUTTONDOWN
        If Not (TabStripStyleCache And TCS_FOCUSNEVER) = TCS_FOCUSNEVER Then
            If (TabStripStyleCache And TCS_FOCUSONBUTTONDOWN) = TCS_FOCUSONBUTTONDOWN Then
                If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            Else
                If GetFocus() <> hWnd Then SetFocusAPI UserControl.hWnd ' UCNoSetFocusFwd not applicable
            End If
        End If
    Case WM_MBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    Case WM_SETCURSOR
        If LoWord(CLng(lParam)) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(NULL_PTR, MousePointerID(PropMousePointer))
                WindowProcControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_ERASEBKGND
        If PropDoubleBuffer = True And (TabStripDoubleBufferEraseBkgDC <> wParam Or TabStripDoubleBufferEraseBkgDC = NULL_PTR) And WindowFromDC(wParam) = hWnd Then
            WindowProcControl = 0
        Else
            Dim ClientRect1 As RECT
            GetClientRect hWnd, ClientRect1
            FillRect wParam, ClientRect1, GetSysColorBrush(COLOR_BTNFACE)
            If PropTransparent = True Then
                If TabStripTransparentBrush = NULL_PTR Then TabStripTransparentBrush = CreateTransparentBrush(wParam)
            End If
            If TabStripBackColorBrush <> NULL_PTR Or TabStripTransparentBrush <> NULL_PTR Then
                Dim Count As Long, i As Long, RC As RECT
                Count = CLng(SendMessage(hWnd, TCM_GETITEMCOUNT, 0, ByVal 0&))
                Dim hRgn As LongPtr, hRgnTab As LongPtr, hRgnFill As LongPtr
                hRgn = CreateRectRgn(0, 0, 0, 0)
                Dim Placement As TbsPlacementConstants
                If ComCtlsSupportLevel() = 0 Then Placement = PropPlacement Else Placement = TbsPlacementTop
                If PropStyle = TbsStyleTabs Then
                    ' Calculate and exclude client area for 'tabs' style only.
                    Select Case Placement
                        Case TbsPlacementTop
                            ClientRect1.Bottom = ClientRect1.Top
                        Case TbsPlacementBottom
                            ClientRect1.Top = ClientRect1.Bottom
                        Case TbsPlacementLeft
                            ClientRect1.Right = ClientRect1.Left
                        Case TbsPlacementRight
                            ClientRect1.Left = ClientRect1.Right
                    End Select
                End If
                For i = 1 To Count
                    If SendMessage(hWnd, TCM_GETITEMRECT, i - 1, ByVal VarPtr(RC)) <> 0 Then
                        hRgnTab = CreateRectRgn(RC.Left, RC.Top, RC.Right, RC.Bottom)
                        If hRgnTab <> NULL_PTR Then
                            CombineRgn hRgn, hRgn, hRgnTab, RGN_OR
                            DeleteObject hRgnTab
                            hRgnTab = NULL_PTR
                        End If
                        Select Case Placement
                            Case TbsPlacementTop
                                If RC.Bottom > ClientRect1.Bottom Then ClientRect1.Bottom = RC.Bottom
                            Case TbsPlacementBottom
                                If RC.Top < ClientRect1.Top Then ClientRect1.Top = RC.Top
                            Case TbsPlacementLeft
                                If RC.Right > ClientRect1.Right Then ClientRect1.Right = RC.Right
                            Case TbsPlacementRight
                                If RC.Left < ClientRect1.Left Then ClientRect1.Left = RC.Left
                        End Select
                    End If
                Next i
                hRgnFill = CreateRectRgn(ClientRect1.Left, ClientRect1.Top, ClientRect1.Right, ClientRect1.Bottom)
                CombineRgn hRgnFill, hRgnFill, hRgn, RGN_DIFF
                If TabStripTransparentBrush = NULL_PTR Then
                    FillRgn wParam, hRgnFill, TabStripBackColorBrush
                Else
                    FillRgn wParam, hRgnFill, TabStripTransparentBrush
                End If
                DeleteObject hRgnFill
                DeleteObject hRgn
            End If
            WindowProcControl = 1
        End If
        Exit Function
    Case WM_PAINT
        If wParam = 0 Then
            If PropDoubleBuffer = True Then
                Dim ClientRect2 As RECT, hDC As LongPtr
                Dim hDCBmp As LongPtr
                Dim hBmp As LongPtr, hBmpOld As LongPtr
                GetClientRect hWnd, ClientRect2
                Dim PS As PAINTSTRUCT
                hDC = BeginPaint(hWnd, PS)
                With PS
                hDCBmp = CreateCompatibleDC(hDC)
                If hDCBmp <> NULL_PTR Then
                    hBmp = CreateCompatibleBitmap(hDC, ClientRect2.Right - ClientRect2.Left, ClientRect2.Bottom - ClientRect2.Top)
                    If hBmp <> NULL_PTR Then
                        hBmpOld = SelectObject(hDCBmp, hBmp)
                        TabStripDoubleBufferEraseBkgDC = hDCBmp
                        SendMessage hWnd, WM_PRINT, hDCBmp, ByVal PRF_CLIENT Or PRF_ERASEBKGND
                        TabStripDoubleBufferEraseBkgDC = NULL_PTR
                        With PS.RCPaint
                        BitBlt hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, hDCBmp, .Left, .Top, vbSrcCopy
                        End With
                        SelectObject hDCBmp, hBmpOld
                        DeleteObject hBmp
                    End If
                    DeleteDC hDCBmp
                End If
                End With
                EndPaint hWnd, PS
                WindowProcControl = 0
                Exit Function
            End If
        Else
            SendMessage hWnd, WM_PRINT, wParam, ByVal PRF_CLIENT Or PRF_ERASEBKGND
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_MOUSEWHEEL
        If PropTabScrollWheel = True Then
            Static WheelDelta As Long, LastWheelDelta As Long
            If Sgn(HiWord(CLng(wParam))) <> Sgn(LastWheelDelta) Then WheelDelta = 0
            WheelDelta = WheelDelta + HiWord(CLng(wParam))
            If Abs(WheelDelta) >= 120 Then
                Dim CurrIndex As Long
                CurrIndex = CLng(SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&)) + 1
                If Sgn(WheelDelta) = -1 Then
                    If CurrIndex < Me.Tabs.Count Then Me.Tabs(CurrIndex + 1).Selected = True
                Else
                    If CurrIndex > 1 Then Me.Tabs(CurrIndex - 1).Selected = True
                End If
                WheelDelta = 0
            End If
            LastWheelDelta = HiWord(CLng(wParam))
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = CLng(wParam) And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            TabStripCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If TabStripCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(TabStripCharCodeCache And &HFFFF&)
            TabStripCharCodeCache = 0
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
    Case WM_PARENTNOTIFY
        If LoWord(CLng(wParam)) = WM_CREATE And lParam <> 0 Then
            If PropVisualStyles = True Then
                ActivateVisualStyles lParam
            Else
                RemoveVisualStyles lParam
            End If
        End If
    Case WM_STYLECHANGED
        If wParam = GWL_STYLE Then CopyMemory TabStripStyleCache, ByVal UnsignedAdd(lParam, 4), 4
    
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
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                If TabStripMouseOver = False And PropMouseTrack = True Then
                    TabStripMouseOver = True
                    RaiseEvent MouseEnter
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP
                RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONUP
                RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONUP
                RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
        End Select
    Case WM_MOUSELEAVE
        If TabStripMouseOver = True Then
            TabStripMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = TabStripHandle Then
            Dim Index As Long
            Select Case NM.Code
                Case TCN_SELCHANGING
                    Index = CLng(SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&))
                    If Index >= 0 Then
                        Dim Cancel As Boolean
                        RaiseEvent TabBeforeClick(Me.Tabs(Index + 1), Cancel)
                        If Cancel = True Then
                            WindowProcUserControl = 1
                            Exit Function
                        End If
                    End If
                Case TCN_SELCHANGE
                    Index = CLng(SendMessage(TabStripHandle, TCM_GETCURSEL, 0, ByVal 0&))
                    If Index >= 0 Then RaiseEvent TabClick(Me.Tabs(Index + 1))
            End Select
        ElseIf NM.hWndFrom = TabStripToolTipHandle And TabStripToolTipHandle <> NULL_PTR Then
            Select Case NM.Code
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    With NMTTDI
                    If PropRightToLeft = True And PropRightToLeftLayout = False Then
                        If Not (.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then
                            .uFlags = .uFlags Or TTF_RTLREADING
                            CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                        End If
                    End If
                    Dim Text As String
                    Text = Me.Tabs(.hdr.IDFrom + 1).ToolTipText
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
    Case WM_PRINTCLIENT
        If TabStripHandle <> NULL_PTR Then
            If PropTransparent = True Then
                If TabStripTransparentBrush = NULL_PTR Then TabStripTransparentBrush = CreateTransparentBrush(wParam)
            End If
            If TabStripBackColorBrush <> NULL_PTR Or TabStripTransparentBrush <> NULL_PTR Then
                Dim RC As RECT
                GetClientRect TabStripHandle, RC
                If TabStripTransparentBrush = NULL_PTR Then
                    FillRect wParam, RC, TabStripBackColorBrush
                Else
                    FillRect wParam, RC, TabStripTransparentBrush
                End If
                WindowProcUserControl = 0
                Exit Function
            End If
        End If
    Case WM_DRAWITEM
        Dim DIS As DRAWITEMSTRUCT
        CopyMemory DIS, ByVal lParam, LenB(DIS)
        If DIS.CtlType = ODT_TAB And DIS.hWndItem = TabStripHandle And DIS.ItemID > -1 Then
            With DIS
            #If Win64 Then
            Dim hDC32 As Long
            CopyMemory ByVal VarPtr(hDC32), ByVal VarPtr(.hDC), 4
            RaiseEvent ItemDraw(Me.Tabs(.ItemID + 1), .ItemAction, .ItemState, hDC32, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
            #Else
            RaiseEvent ItemDraw(Me.Tabs(.ItemID + 1), .ItemAction, .ItemState, .hDC, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
            #End If
            End With
            WindowProcUserControl = 1
            Exit Function
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If lParam = NF_QUERY Then
            Const NFR_UNICODE As Long = 2
            Const NFR_ANSI As Long = 1
            WindowProcUserControl = NFR_UNICODE
            Exit Function
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI TabStripHandle
End Function

Private Function WindowProcControlDesignMode(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_ERASEBKGND
        WindowProcControlDesignMode = WindowProcControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function

Private Function WindowProcUserControlDesignMode(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_PRINTCLIENT
        WindowProcUserControlDesignMode = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcUserControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
