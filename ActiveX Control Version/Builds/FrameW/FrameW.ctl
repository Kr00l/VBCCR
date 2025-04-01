VERSION 5.00
Begin VB.UserControl FrameW 
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ControlContainer=   -1  'True
   DrawStyle       =   5  'Transparent
   ForwardFocus    =   -1  'True
   PropertyPages   =   "FrameW.ctx":0000
   ScaleHeight     =   1800
   ScaleWidth      =   2400
   ToolboxBitmap   =   "FrameW.ctx":0035
End
Attribute VB_Name = "FrameW"
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

#Const ImplementThemedButton = True

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
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
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
Private Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As LongPtr, ByVal lpchText As LongPtr, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare PtrSafe Function ExcludeClipRect Lib "gdi32" (ByVal hDC As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function SelectClipRgn Lib "gdi32" (ByVal hDC As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hWndFrom As LongPtr, ByVal hWndTo As LongPtr, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function DrawEdge Lib "user32" (ByVal hDC As LongPtr, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function RevokeDragDrop Lib "ole32" (ByVal hWnd As LongPtr) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RevokeDragDrop Lib "ole32" (ByVal hWnd As Long) As Long
#End If

#If ImplementThemedButton = True Then

Private Enum UxThemeButtonParts
BP_PUSHBUTTON = 1
BP_RADIOBUTTON = 2
BP_CHECKBOX = 3
BP_GROUPBOX = 4
BP_USERBUTTON = 5
End Enum
Private Enum UxThemeGroupBoxStates
GBS_NORMAL = 1
GBS_DISABLED = 2
End Enum
Private Const DTT_TEXTCOLOR As Long = 1
Private Type DTTOPTS
dwSize As Long
dwFlags As Long
crText As Long
crBorder As Long
crShadow As Long
eTextShadowType As Long
PTShadowOffset As POINTAPI
iBorderSize As Long
iFontPropId As Long
iColorPropId As Long
iStateId As Long
fApplyOverlay As Long
iGlowSize As Long
End Type
#If VBA7 Then
Private Declare PtrSafe Function GetThemeTextExtent Lib "uxtheme" (ByVal Theme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As LongPtr, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByRef lpBoundingRect As RECT, ByRef lpExtentRect As RECT) As Long
Private Declare PtrSafe Function DrawThemeText Lib "uxtheme" (ByVal Theme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As LongPtr, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, ByRef pRect As RECT) As Long
Private Declare PtrSafe Function DrawThemeTextEx Lib "uxtheme" (ByVal Theme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As LongPtr, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByRef lpRect As RECT, ByRef lpOptions As DTTOPTS) As Long
Private Declare PtrSafe Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare PtrSafe Function OpenThemeData Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal lpszClassList As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseThemeData Lib "uxtheme" (ByVal Theme As LongPtr) As Long
#Else
Private Declare Function GetThemeTextExtent Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByRef lpBoundingRect As RECT, ByRef lpExtentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawThemeTextEx Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByRef lpRect As RECT, ByRef lpOptions As DTTOPTS) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal lpszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long
#End If

#End If

Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
#If VBA7 Then
Private Const HWND_DESKTOP As LongPtr = &H0
#Else
Private Const HWND_DESKTOP As Long = &H0
#End If
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_NOCLIP As Long = &H100
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_RTLREADING As Long = &H20000
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const BF_LEFT As Long = 1
Private Const BF_TOP As Long = 2
Private Const BF_RIGHT As Long = 4
Private Const BF_BOTTOM As Long = 8
Private Const BF_RECT As Long = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Private Const BF_MONO As Long = &H8000&
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private FrameMouseOver As Boolean
Private FrameDesignMode As Boolean
Private FramePictureRenderFlag As Integer
Private DispIdBorderStyle As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBorderStyle As Integer
Private PropCaption As String
Private PropUseMnemonic As Boolean
Private PropAlignment As VBRUN.AlignmentConstants
Private PropTransparent As Boolean
Private PropPicture As IPictureDisp
Private PropPictureAlignment As CCLeftRightAlignmentConstants

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispId As Long, ByRef DisplayName As String)
If DispId = DispIdBorderStyle Then
    Select Case PropBorderStyle
        Case vbBSNone: DisplayName = vbBSNone & " - None"
        Case vbFixedSingle: DisplayName = vbFixedSingle & " - Fixed Single"
    End Select
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispId As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispId = DispIdBorderStyle Then
    ReDim StringsOut(0 To (1 + 1)) As String
    ReDim CookiesOut(0 To (1 + 1)) As Long
    StringsOut(0) = vbBSNone & " - None": CookiesOut(0) = vbBSNone
    StringsOut(1) = vbFixedSingle & " - Fixed Single": CookiesOut(1) = vbFixedSingle
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispId As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispId = DispIdBorderStyle Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIdBorderStyle = 0 Then DispIdBorderStyle = GetDispId(Me, "BorderStyle")
On Error Resume Next
FrameDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set Me.Font = Ambient.Font
PropVisualStyles = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = vbFixedSingle
PropCaption = Ambient.DisplayName
PropUseMnemonic = True
If PropRightToLeft = False Then PropAlignment = vbLeftJustify Else PropAlignment = vbRightJustify
PropTransparent = False
Set PropPicture = Nothing
If PropRightToLeft = False Then PropPictureAlignment = CCLeftRightAlignmentLeft Else PropPictureAlignment = CCLeftRightAlignmentRight
If FrameDesignMode = False Then Call ComCtlsSetSubclass(UserControl.hWnd, Me, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIdBorderStyle = 0 Then DispIdBorderStyle = GetDispId(Me, "BorderStyle")
On Error Resume Next
FrameDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set Me.Font = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Appearance = .ReadProperty("Appearance", CCAppearance3D)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
Me.MousePointer = .ReadProperty("MousePointer", 0)
Set Me.MouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = .ReadProperty("BorderStyle", vbFixedSingle)
PropCaption = .ReadProperty("Caption", vbNullString) ' Unicode not necessary
PropUseMnemonic = .ReadProperty("UseMnemonic", True)
PropAlignment = .ReadProperty("Alignment", vbLeftJustify)
PropTransparent = .ReadProperty("Transparent", False)
Set PropPicture = .ReadProperty("Picture", Nothing)
PropPictureAlignment = .ReadProperty("PictureAlignment", CCLeftRightAlignmentLeft)
End With
If PropUseMnemonic = True Then
    UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
Else
    UserControl.AccessKeys = vbNullString
End If
If FrameDesignMode = False Then Call ComCtlsSetSubclass(UserControl.hWnd, Me, 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Appearance", Me.Appearance, CCAppearance3D
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", Me.MouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BorderStyle", PropBorderStyle, vbFixedSingle
.WriteProperty "Caption", PropCaption, vbNullString ' Unicode not necessary
.WriteProperty "UseMnemonic", PropUseMnemonic, True
.WriteProperty "Alignment", PropAlignment, vbLeftJustify
.WriteProperty "Transparent", PropTransparent, False
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "PictureAlignment", PropPictureAlignment, CCLeftRightAlignmentLeft
End With
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FrameMouseOver = False And PropMouseTrack = True Then
    FrameMouseOver = True
    RaiseEvent MouseEnter
    Call ComCtlsRequestMouseLeave(UserControl.hWnd)
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
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
Static PrevHeight As Long, PrevWidth As Long
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
Call DrawFrame
InProc = False
If PrevHeight <> .ScaleHeight Or PrevWidth <> .ScaleWidth Then
    PrevHeight = .ScaleHeight
    PrevWidth = .ScaleWidth
    RaiseEvent Resize
End If
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
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
hWnd = UserControl.hWnd
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
Set PropFont = NewFont
Set UserControl.Font = PropFont
Call DrawFrame
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Set UserControl.Font = PropFont
Call DrawFrame
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
Call DrawFrame
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Appearance() As CCAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets a value that determines whether an object is painted two-dimensional or with 3-D effects."
Attribute Appearance.VB_UserMemId = -520
Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal Value As CCAppearanceConstants)
Select Case Value
    Case CCAppearanceFlat, CCAppearance3D
        UserControl.Appearance = Value
    Case Else
        Err.Raise 380
End Select
UserControl.ForeColor = IIf(UserControl.Appearance = CCAppearanceFlat, vbWindowText, vbButtonText)
Call DrawFrame
UserControl.PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
Call DrawFrame
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
UserControl.ForeColor = Value
Call DrawFrame
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
Call DrawFrame
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDropMode() As OLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
' Setting OLEDropMode to OLEDropModeManual will fail when windowless controls are contained in the user control.
Const DRAGDROP_E_ALREADYREGISTERED As Long = &H80040101
Select Case Value
    Case OLEDropModeNone, OLEDropModeManual
        On Error Resume Next
        UserControl.OLEDropMode = Value
        If Err.Number = DRAGDROP_E_ALREADYREGISTERED Then
            RevokeDragDrop UserControl.hWnd
            UserControl.OLEDropMode = Value
        End If
        On Error GoTo 0
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
If FrameDesignMode = False Then
    Select Case PropMousePointer
        Case vbIconPointer, 16
            UserControl.MousePointer = vbDefault
        Case Else
            UserControl.MousePointer = PropMousePointer
    End Select
End If
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set UserControl.MouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = NULL_PTR Then
        Set UserControl.MouseIcon = Value
    Else
        If FrameDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
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
If PropRightToLeft = False Then
    If PropAlignment = vbRightJustify Then PropAlignment = vbLeftJustify
    If PropPictureAlignment = CCLeftRightAlignmentRight Then PropPictureAlignment = CCLeftRightAlignmentLeft
Else
    If PropAlignment = vbLeftJustify Then PropAlignment = vbRightJustify
    If PropPictureAlignment = CCLeftRightAlignmentLeft Then PropPictureAlignment = CCLeftRightAlignmentRight
End If
Call DrawFrame
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

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As Integer)
Select Case Value
    Case vbBSNone, vbFixedSingle
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
Call DrawFrame
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
Caption = PropCaption
End Property

Public Property Let Caption(ByVal Value As String)
If PropCaption = Value Then Exit Property
PropCaption = Value
If PropUseMnemonic = True Then UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
Call DrawFrame
UserControl.PropertyChanged "Caption"
End Property

Public Property Get UseMnemonic() As Boolean
Attribute UseMnemonic.VB_Description = "Returns/sets a value that specifies whether an & in the caption property defines an access key."
UseMnemonic = PropUseMnemonic
End Property

Public Property Let UseMnemonic(ByVal Value As Boolean)
PropUseMnemonic = Value
If PropUseMnemonic = True Then
    UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
Else
    UserControl.AccessKeys = vbNullString
End If
Call DrawFrame
UserControl.PropertyChanged "UseMnemonic"
End Property

Public Property Get Alignment() As VBRUN.AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As VBRUN.AlignmentConstants)
Select Case Value
    Case vbLeftJustify, vbCenter, vbRightJustify
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
Call DrawFrame
UserControl.PropertyChanged "Alignment"
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value indicating if the background is a replica of the underlying background to simulate transparency."
Transparent = PropTransparent
End Property

Public Property Let Transparent(ByVal Value As Boolean)
PropTransparent = Value
Call DrawFrame
UserControl.PropertyChanged "Transparent"
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropPicture = Nothing
Else
    Set UserControl.Picture = Value
    Set PropPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
End If
FramePictureRenderFlag = 0
Call DrawFrame
UserControl.PropertyChanged "Picture"
End Property

Public Property Get PictureAlignment() As CCLeftRightAlignmentConstants
Attribute PictureAlignment.VB_Description = "Returns/sets the picture alignment."
PictureAlignment = PropPictureAlignment
End Property

Public Property Let PictureAlignment(ByVal Value As CCLeftRightAlignmentConstants)
Select Case Value
    Case CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
        PropPictureAlignment = Value
    Case Else
        Err.Raise 380
End Select
Call DrawFrame
UserControl.PropertyChanged "PictureAlignment"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
Call DrawFrame
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get ContainedControls() As VBRUN.ContainedControls
Attribute ContainedControls.VB_Description = "Returns a collection that allows access to the controls contained within the control that were added to the control by the developer who uses the control."
Set ContainedControls = UserControl.ContainedControls
End Property

Private Sub DrawFrame()
With UserControl
.Cls
.AutoRedraw = True
Set .Picture = Nothing
If PropTransparent = True Then
    Dim WndRect As RECT, P As POINTAPI
    GetWindowRect .hWnd, WndRect
    MapWindowPoints HWND_DESKTOP, GetParent(.hWnd), WndRect, 2
    P.X = WndRect.Left
    P.Y = WndRect.Top
    SetViewportOrgEx .hDC, -P.X, -P.Y, P
    SendMessage GetParent(.hWnd), WM_PAINT, .hDC, ByVal 0&
    SetViewportOrgEx .hDC, P.X, P.Y, P
End If
If PropBorderStyle <> vbBSNone Then
    Dim ClientRect As RECT, BoundingRect As RECT, ExtentRect As RECT, DrawFlags As Long, OldBkMode As Long
    Dim TextRect As RECT, CX As Long
    GetClientRect .hWnd, ClientRect
    LSet BoundingRect = ClientRect
    BoundingRect.Left = 9
    BoundingRect.Right = BoundingRect.Right - BoundingRect.Left
    DrawFlags = DT_NOCLIP Or DT_SINGLELINE
    If PropRightToLeft = True Then DrawFlags = DrawFlags Or DT_RTLREADING
    If PropUseMnemonic = False Then DrawFlags = DrawFlags Or DT_NOPREFIX
    Select Case PropAlignment
        Case vbLeftJustify
            DrawFlags = DrawFlags Or DT_LEFT
        Case vbCenter
            DrawFlags = DrawFlags Or DT_CENTER
        Case vbRightJustify
            DrawFlags = DrawFlags Or DT_RIGHT
    End Select
    OldBkMode = SetBkMode(.hDC, 1)
    Dim PictureWidth As Long, PictureHeight As Long
    Dim PictureLeft As Long, PictureTop As Long
    If Not PropPicture Is Nothing Then
        If PropPicture.Handle <> NULL_PTR Then
            PictureWidth = CHimetricToPixel_X(PropPicture.Width)
            PictureHeight = CHimetricToPixel_Y(PropPicture.Height)
            PictureTop = BoundingRect.Top
        End If
    End If
    Dim Theme As LongPtr
    
    #If ImplementThemedButton = True Then
    
    If EnabledVisualStyles() = True And PropVisualStyles = True Then Theme = OpenThemeData(.hWnd, StrPtr("Button"))
    If Theme <> NULL_PTR Then
        Dim ButtonPart As Long, GroupBoxState As Long
        ButtonPart = BP_GROUPBOX
        If .Enabled = True Then
            GroupBoxState = GBS_NORMAL
        Else
            GroupBoxState = GBS_DISABLED
        End If
        GetThemeTextExtent Theme, .hDC, ButtonPart, GroupBoxState, StrPtr("A"), 1, DrawFlags, BoundingRect, ExtentRect
        If PictureHeight <= (ExtentRect.Bottom - ExtentRect.Top) Then
            ClientRect.Top = ClientRect.Top + ((ExtentRect.Bottom - ExtentRect.Top) \ 2)
            PictureTop = PictureTop + (((ExtentRect.Bottom - ExtentRect.Top) - PictureHeight) \ 2)
        Else
            ClientRect.Top = ClientRect.Top + (PictureHeight \ 2)
            BoundingRect.Top = BoundingRect.Top + ((PictureHeight - (ExtentRect.Bottom - ExtentRect.Top)) \ 2)
        End If
        If Not PropCaption = vbNullString Then
            GetThemeTextExtent Theme, .hDC, ButtonPart, GroupBoxState, StrPtr(PropCaption), Len(PropCaption), DrawFlags, BoundingRect, ExtentRect
            LSet TextRect = BoundingRect
            If PictureWidth > 0 And PictureHeight > 0 Then
                Select Case PropAlignment
                    Case vbLeftJustify
                        If PropPictureAlignment = CCLeftRightAlignmentLeft Then TextRect.Left = TextRect.Left + PictureWidth + 2
                    Case vbCenter
                        If PropPictureAlignment = CCLeftRightAlignmentLeft Then
                            TextRect.Left = TextRect.Left + PictureWidth + 2
                        ElseIf PropPictureAlignment = CCLeftRightAlignmentRight Then
                            TextRect.Left = TextRect.Left - PictureWidth - 2
                        End If
                    Case vbRightJustify
                        If PropPictureAlignment = CCLeftRightAlignmentRight Then TextRect.Right = TextRect.Right - PictureWidth - 2
                End Select
            End If
            If ComCtlsSupportLevel() >= 2 Then
                Dim DTTO As DTTOPTS
                DTTO.dwSize = LenB(DTTO)
                DTTO.dwFlags = DTT_TEXTCOLOR
                If .Enabled = True Then
                    DTTO.crText = WinColor(.ForeColor)
                Else
                    DTTO.crText = WinColor(vbGrayText)
                End If
                DrawThemeTextEx Theme, .hDC, ButtonPart, GroupBoxState, StrPtr(PropCaption), Len(PropCaption), DrawFlags, TextRect, DTTO
            Else
                DrawThemeText Theme, .hDC, ButtonPart, GroupBoxState, StrPtr(PropCaption), Len(PropCaption), DrawFlags, 0, TextRect
            End If
            CX = (BoundingRect.Right - BoundingRect.Left) - (ExtentRect.Right - ExtentRect.Left)
            Select Case PropAlignment
                Case vbCenter
                    ExtentRect.Left = ExtentRect.Left + (CX \ 2)
                    ExtentRect.Right = ExtentRect.Right + (CX \ 2)
                Case vbRightJustify
                    ExtentRect.Left = ExtentRect.Left + CX
                    ExtentRect.Right = ExtentRect.Right + CX
            End Select
            If PictureWidth > 0 And PictureHeight > 0 Then
                Select Case PropAlignment
                    Case vbLeftJustify
                        ExtentRect.Right = ExtentRect.Right + PictureWidth + 2
                    Case vbCenter
                        ExtentRect.Left = ExtentRect.Left - ((PictureWidth + 2) \ 2)
                        ExtentRect.Right = ExtentRect.Right + ((PictureWidth + 2) \ 2)
                    Case vbRightJustify
                        ExtentRect.Left = ExtentRect.Left - PictureWidth - 2
                End Select
                If PictureHeight > ExtentRect.Bottom Then ExtentRect.Bottom = PictureHeight
                If PropPictureAlignment = CCLeftRightAlignmentLeft Then
                    PictureLeft = ExtentRect.Left
                Else
                    PictureLeft = ExtentRect.Right - PictureWidth
                End If
                Call RenderPicture(PropPicture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, FramePictureRenderFlag)
            End If
            ExcludeClipRect .hDC, ExtentRect.Left - 2, ExtentRect.Top, ExtentRect.Right + 2, ExtentRect.Bottom
        ElseIf PictureWidth > 0 And PictureHeight > 0 Then
            ExtentRect.Top = PictureTop
            ExtentRect.Bottom = ExtentRect.Top + PictureHeight
            Select Case PropAlignment
                Case vbLeftJustify
                    ExtentRect.Left = BoundingRect.Left
                    ExtentRect.Right = ExtentRect.Left + PictureWidth
                Case vbCenter
                    ExtentRect.Left = BoundingRect.Left + ((BoundingRect.Right - BoundingRect.Left) \ 2) - (PictureWidth \ 2)
                    ExtentRect.Right = ExtentRect.Left + PictureWidth
                Case vbRightJustify
                    ExtentRect.Left = BoundingRect.Right - PictureWidth
                    ExtentRect.Right = BoundingRect.Right
            End Select
            PictureLeft = ExtentRect.Left
            Call RenderPicture(PropPicture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, FramePictureRenderFlag)
            ExcludeClipRect .hDC, ExtentRect.Left - 2, ExtentRect.Top, ExtentRect.Right + 2, ExtentRect.Bottom
        End If
        DrawThemeBackground Theme, .hDC, ButtonPart, GroupBoxState, ClientRect, ClientRect
        SelectClipRgn .hDC, NULL_PTR
        CloseThemeData Theme
    End If
    
    #End If
    
    If Theme = NULL_PTR Then
        Dim Size As SIZEAPI
        GetTextExtentPoint32 .hDC, ByVal StrPtr("A"), 1, Size
        If PictureHeight <= Size.CY Then
            ClientRect.Top = ClientRect.Top + (Size.CY \ 2)
            PictureTop = PictureTop + ((Size.CY - PictureHeight) \ 2)
        Else
            ClientRect.Top = ClientRect.Top + (PictureHeight \ 2)
            BoundingRect.Top = BoundingRect.Top + ((PictureHeight - Size.CY) \ 2)
        End If
        If Not PropCaption = vbNullString Then
            GetTextExtentPoint32 .hDC, ByVal StrPtr(PropCaption), Len(PropCaption), Size
            LSet ExtentRect = BoundingRect
            ExtentRect.Right = ExtentRect.Left + Size.CX
            ExtentRect.Bottom = ExtentRect.Top + Size.CY
            LSet TextRect = BoundingRect
            If PictureWidth > 0 And PictureHeight > 0 Then
                Select Case PropAlignment
                    Case vbLeftJustify
                        If PropPictureAlignment = CCLeftRightAlignmentLeft Then TextRect.Left = TextRect.Left + PictureWidth + 2
                    Case vbCenter
                        If PropPictureAlignment = CCLeftRightAlignmentLeft Then
                            TextRect.Left = TextRect.Left + PictureWidth + 2
                        ElseIf PropPictureAlignment = CCLeftRightAlignmentRight Then
                            TextRect.Left = TextRect.Left - PictureWidth - 2
                        End If
                    Case vbRightJustify
                        If PropPictureAlignment = CCLeftRightAlignmentRight Then TextRect.Right = TextRect.Right - PictureWidth - 2
                End Select
            End If
            Dim OldTextColor As Long
            If .Enabled = True Then
                OldTextColor = SetTextColor(.hDC, WinColor(.ForeColor))
            Else
                OldTextColor = SetTextColor(.hDC, WinColor(vbGrayText))
            End If
            DrawText .hDC, StrPtr(PropCaption), Len(PropCaption), TextRect, DrawFlags
            SetTextColor .hDC, OldTextColor
            CX = (BoundingRect.Right - BoundingRect.Left) - (ExtentRect.Right - ExtentRect.Left)
            Select Case PropAlignment
                Case vbCenter
                    ExtentRect.Left = ExtentRect.Left + (CX \ 2)
                    ExtentRect.Right = ExtentRect.Right + (CX \ 2)
                Case vbRightJustify
                    ExtentRect.Left = ExtentRect.Left + CX
                    ExtentRect.Right = ExtentRect.Right + CX
            End Select
            If PictureWidth > 0 And PictureHeight > 0 Then
                Select Case PropAlignment
                    Case vbLeftJustify
                        ExtentRect.Right = ExtentRect.Right + PictureWidth + 2
                    Case vbCenter
                        ExtentRect.Left = ExtentRect.Left - ((PictureWidth + 2) \ 2)
                        ExtentRect.Right = ExtentRect.Right + ((PictureWidth + 2) \ 2)
                    Case vbRightJustify
                        ExtentRect.Left = ExtentRect.Left - PictureWidth - 2
                End Select
                If PictureHeight > ExtentRect.Bottom Then ExtentRect.Bottom = PictureHeight
                If PropPictureAlignment = CCLeftRightAlignmentLeft Then
                    PictureLeft = ExtentRect.Left
                Else
                    PictureLeft = ExtentRect.Right - PictureWidth
                End If
                Call RenderPicture(PropPicture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, FramePictureRenderFlag)
            End If
            ExcludeClipRect .hDC, ExtentRect.Left - 2, ExtentRect.Top, ExtentRect.Right + 2, ExtentRect.Bottom
        ElseIf PictureWidth > 0 And PictureHeight > 0 Then
            ExtentRect.Top = PictureTop
            ExtentRect.Bottom = ExtentRect.Top + PictureHeight
            Select Case PropAlignment
                Case vbLeftJustify
                    ExtentRect.Left = BoundingRect.Left
                    ExtentRect.Right = ExtentRect.Left + PictureWidth
                Case vbCenter
                    ExtentRect.Left = BoundingRect.Left + ((BoundingRect.Right - BoundingRect.Left) \ 2) - (PictureWidth \ 2)
                    ExtentRect.Right = ExtentRect.Left + PictureWidth
                Case vbRightJustify
                    ExtentRect.Left = BoundingRect.Right - PictureWidth
                    ExtentRect.Right = BoundingRect.Right
            End Select
            PictureLeft = ExtentRect.Left
            Call RenderPicture(PropPicture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, FramePictureRenderFlag)
            ExcludeClipRect .hDC, ExtentRect.Left - 2, ExtentRect.Top, ExtentRect.Right + 2, ExtentRect.Bottom
        End If
        DrawEdge .hDC, ClientRect, EDGE_ETCHED, BF_RECT Or IIf(.Appearance = CCAppearanceFlat, BF_MONO, 0)
        SelectClipRgn .hDC, NULL_PTR
    End If
    SetBkMode .hDC, OldBkMode
End If
Set .Picture = .Image
.AutoRedraw = False
End With
End Sub

#If VBA7 Then
Private Function ISubclass_Message(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
#End If
ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_PRINTCLIENT
        Dim ClientRect As RECT
        GetClientRect UserControl.hWnd, ClientRect
        BitBlt wParam, 0, 0, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top, UserControl.hDC, 0, 0, vbSrcCopy
        WindowProcUserControl = 0
        Exit Function
    Case WM_GETTEXTLENGTH
        WindowProcUserControl = Len(PropCaption)
        Exit Function
    Case WM_GETTEXT, WM_SETTEXT
        Dim Length As Long, Text As String
        If wMsg = WM_GETTEXT Then
            If wParam > 0 And lParam <> 0 Then
                Length = Len(PropCaption) + 1
                If wParam < Length Then Length = CLng(wParam)
                Text = Left$(PropCaption, Length - 1) & vbNullChar
                CopyMemory ByVal lParam, ByVal StrPtr(Text), Length * 2
                WindowProcUserControl = Length - 1
            Else
                WindowProcUserControl = 0
            End If
        ElseIf wMsg = WM_SETTEXT Then
            If lParam <> 0 Then Length = lstrlen(lParam)
            If Length > 0 Then
                Text = String$(Length, vbNullChar)
                CopyMemory ByVal StrPtr(Text), ByVal lParam, Length * 2
                Me.Caption = Text
                WindowProcUserControl = 1
            ElseIf lParam = 0 Then
                Me.Caption = vbNullString
                WindowProcUserControl = 1
            Else
                WindowProcUserControl = 0
            End If
        End If
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_MOUSELEAVE Then
    If FrameMouseOver = True Then
        FrameMouseOver = False
        RaiseEvent MouseLeave
    End If
End If
End Function
