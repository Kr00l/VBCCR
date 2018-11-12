VERSION 5.00
Begin VB.UserControl IPAddress 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   HasDC           =   0   'False
   PropertyPages   =   "IPAddress.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "IPAddress.ctx":0035
End
Attribute VB_Name = "IPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As Long
dwHoverTime As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMIPADDRESS
hdr As NMHDR
iField As Long
iValue As Long
End Type
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event FieldChange(ByVal Field As Long)
Attribute FieldChange.VB_Description = "Occurs when the contents of a specific field have changed."
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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Const ICC_INTERNET_CLASSES As Long = &H800
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SW_HIDE As Long = &H0
Private Const TME_LEAVE As Long = &H2, TME_NONCLIENT As Long = &H10
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4, HTBORDER As Long = 18
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_COMMAND As Long = &H111
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
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_NCMOUSELEAVE As Long = &H2A2
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SETFONT As Long = &H30
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_PASTE As Long = &H302
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_USER As Long = &H400
Private Const IPM_CLEARADDRESS As Long = (WM_USER + 100)
Private Const IPM_SETADDRESS As Long = (WM_USER + 101)
Private Const IPM_GETADDRESS As Long = (WM_USER + 102)
Private Const IPM_SETRANGE As Long = (WM_USER + 103)
Private Const IPM_SETFOCUS As Long = (WM_USER + 104)
Private Const IPM_ISBLANK As Long = (WM_USER + 105)
Private Const IPN_FIRST As Long = (-860)
Private Const IPN_FIELDCHANGED As Long = (IPN_FIRST - 0)
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private IPAddressHandle As Long, IPAddressEditHandle(1 To 4) As Long
Private IPAddressFontHandle As Long
Private IPAddressCharCodeCache As Long
Private IPAddressMouseOver(0 To 5) As Boolean
Private IPAddressDesignMode As Boolean, IPAddressTopDesignMode As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropMin(1 To 4) As Byte, PropMax(1 To 4) As Byte
Private PropForeColor As OLE_COLOR

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = wParam And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyTab, vbKeyReturn, vbKeyEscape
            Dim hWnd As Long
            hWnd = GetFocus()
            If hWnd <> 0 Then
                Select Case hWnd
                    Case IPAddressHandle, IPAddressEditHandle(1), IPAddressEditHandle(2), IPAddressEditHandle(3), IPAddressEditHandle(4)
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
                End Select
            End If
    End Select
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropMousePointer, DisplayName)
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_INTERNET_CLASSES)
Call SetVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
IPAddressDesignMode = Not Ambient.UserMode
IPAddressTopDesignMode = Not GetTopUserControl(Me).Ambient.UserMode
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropForeColor = vbWindowText
Call CreateIPAddress
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
IPAddressDesignMode = Not Ambient.UserMode
IPAddressTopDesignMode = Not GetTopUserControl(Me).Ambient.UserMode
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropForeColor = .ReadProperty("ForeColor", vbWindowText)
End With
Call CreateIPAddress
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
.WriteProperty "ForeColor", PropForeColor, vbWindowText
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
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
If DPICorrectionFactor() <> 1 Then
    .Extender.Move .Extender.Left + .ScaleX(1, vbPixels, vbContainerPosition), .Extender.Top + .ScaleY(1, vbPixels, vbContainerPosition)
    .Extender.Move .Extender.Left - .ScaleX(1, vbPixels, vbContainerPosition), .Extender.Top - .ScaleY(1, vbPixels, vbContainerPosition)
End If
If IPAddressHandle <> 0 Then
    Dim RC As RECT
    GetWindowRect IPAddressHandle, RC
    If (RC.Right - RC.Left) <> .ScaleWidth Or (RC.Bottom - RC.Top) <> .ScaleHeight Then Call ReCreateIPAddress
End If
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyIPAddress
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

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = IPAddressHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndEdit(ByVal Field As Long) As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
If IPAddressHandle <> 0 Then hWndEdit = IPAddressEditHandle(Field)
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
Dim OldFontHandle As Long
Set PropFont = NewFont
OldFontHandle = IPAddressFontHandle
IPAddressFontHandle = CreateGDIFontFromOLEFont(PropFont)
If IPAddressHandle <> 0 Then SendMessage IPAddressHandle, WM_SETFONT, IPAddressFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = IPAddressFontHandle
IPAddressFontHandle = CreateGDIFontFromOLEFont(PropFont)
If IPAddressHandle <> 0 Then SendMessage IPAddressHandle, WM_SETFONT, IPAddressFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If IPAddressHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles IPAddressHandle
    Else
        RemoveVisualStyles IPAddressHandle
    End If
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
If IPAddressHandle <> 0 Then EnableWindow IPAddressHandle, IIf(Value = True, 1, 0)
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

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
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
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropMouseIcon = Value
    Else
        If IPAddressDesignMode = True Then
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

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object. Only applicable if the enabled property is set to true. This property is ignored at design time."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
Me.Refresh
UserControl.PropertyChanged "ForeColor"
End Property

Private Sub CreateIPAddress()
If IPAddressHandle <> 0 Then Exit Sub
Dim dwStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
IPAddressHandle = CreateWindowEx(0, StrPtr("SysIPAddress32"), StrPtr("IP Address"), dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If IPAddressHandle <> 0 Then
    Dim hWnd(1 To 4) As Long
    hWnd(1) = FindWindowEx(IPAddressHandle, 0, StrPtr("Edit"), 0)
    hWnd(2) = FindWindowEx(IPAddressHandle, hWnd(1), StrPtr("Edit"), 0)
    hWnd(3) = FindWindowEx(IPAddressHandle, hWnd(2), StrPtr("Edit"), 0)
    hWnd(4) = FindWindowEx(IPAddressHandle, hWnd(3), StrPtr("Edit"), 0)
    IPAddressEditHandle(1) = hWnd(4)
    IPAddressEditHandle(2) = hWnd(3)
    IPAddressEditHandle(3) = hWnd(2)
    IPAddressEditHandle(4) = hWnd(1)
    If IPAddressEditHandle(1) = 0 Or IPAddressEditHandle(2) = 0 Or IPAddressEditHandle(3) = 0 Or IPAddressEditHandle(4) = 0 Then
        ShowWindow IPAddressHandle, SW_HIDE
        SetParent IPAddressHandle, 0
        DestroyWindow IPAddressHandle
        IPAddressHandle = 0
        Erase IPAddressEditHandle()
        Exit Sub
    End If
End If
PropMin(1) = 0
PropMin(2) = 0
PropMin(3) = 0
PropMin(4) = 0
PropMax(1) = 255
PropMax(2) = 255
PropMax(3) = 255
PropMax(4) = 255
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If IPAddressDesignMode = False Then
    If IPAddressHandle <> 0 Then
        Call ComCtlsSetSubclass(IPAddressHandle, Me, 1)
        If IPAddressEditHandle(1) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(1), Me, 2)
        If IPAddressEditHandle(2) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(2), Me, 2)
        If IPAddressEditHandle(3) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(3), Me, 2)
        If IPAddressEditHandle(4) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(4), Me, 2)
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
End If
End Sub

Private Sub ReCreateIPAddress()
If IPAddressDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim FieldText(1 To 4) As String, Field As Long
    If IPAddressHandle <> 0 Then
        For Field = 1 To 4
            FieldText(Field) = String(SendMessage(IPAddressEditHandle(Field), WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
            SendMessage IPAddressEditHandle(Field), WM_GETTEXT, Len(FieldText(Field)) + 1, ByVal StrPtr(FieldText(Field))
        Next Field
    End If
    Call DestroyIPAddress
    Call CreateIPAddress
    If IPAddressHandle <> 0 Then
        For Field = 1 To 4
            If Not FieldText(Field) = vbNullString Then SendMessage IPAddressEditHandle(Field), WM_SETTEXT, 0, ByVal StrPtr(FieldText(Field))
        Next Field
    End If
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    Call DestroyIPAddress
    Call CreateIPAddress
End If
End Sub

Private Sub DestroyIPAddress()
If IPAddressHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(IPAddressHandle)
Call ComCtlsRemoveSubclass(IPAddressEditHandle(1))
Call ComCtlsRemoveSubclass(IPAddressEditHandle(2))
Call ComCtlsRemoveSubclass(IPAddressEditHandle(3))
Call ComCtlsRemoveSubclass(IPAddressEditHandle(4))
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow IPAddressHandle, SW_HIDE
SetParent IPAddressHandle, 0
DestroyWindow IPAddressHandle
IPAddressHandle = 0
Erase IPAddressEditHandle()
If IPAddressFontHandle <> 0 Then
    DeleteObject IPAddressFontHandle
    IPAddressFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get Min(ByVal Field As Long) As Byte
Attribute Min.VB_Description = "Returns/sets the minimum value that the specified field accepts."
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
Min = PropMin(Field)
End Property

Public Property Let Min(ByVal Field As Long, ByVal Value As Byte)
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
If Value > PropMax(Field) Then Value = PropMax(Field)
PropMin(Field) = Value
If IPAddressHandle <> 0 Then SendMessage IPAddressHandle, IPM_SETRANGE, Field - 1, ByVal MakeWord(PropMin(Field), PropMax(Field))
End Property

Public Property Get Max(ByVal Field As Long) As Byte
Attribute Max.VB_Description = "Returns/sets the maximum value that the specified field accepts."
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
Max = PropMax(Field)
End Property

Public Property Let Max(ByVal Field As Long, ByVal Value As Byte)
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
If Value < PropMin(Field) Then Value = PropMin(Field)
PropMax(Field) = Value
If IPAddressHandle <> 0 Then SendMessage IPAddressHandle, IPM_SETRANGE, Field - 1, ByVal MakeWord(PropMin(Field), PropMax(Field))
End Property

Public Property Get FieldValue(ByVal Field As Long) As Byte
Attribute FieldValue.VB_Description = "Returns/sets the byte value of the specified field."
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
If IPAddressHandle <> 0 Then
    Dim Text As String
    Text = String(SendMessage(IPAddressEditHandle(Field), WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage IPAddressEditHandle(Field), WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
    If Not Text = vbNullString Then
        On Error Resume Next
        FieldValue = CByte(Text)
        On Error GoTo 0
    End If
End If
End Property

Public Property Let FieldValue(ByVal Field As Long, ByVal Value As Byte)
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
If IPAddressHandle <> 0 Then
    If Value < PropMin(Field) Then Value = PropMin(Field)
    If Value > PropMax(Field) Then Value = PropMax(Field)
    Dim Text As String
    Text = CStr(Value)
    SendMessage IPAddressEditHandle(Field), WM_SETTEXT, 0, ByVal StrPtr(Text)
End If
End Property

Public Property Get Address() As String
Attribute Address.VB_Description = "Returns/sets the currently displayed IP address. Setting this property to an empty string clears the displayed IP address. If this property is an empty string then the displayed IP address is blank."
Attribute Address.VB_UserMemId = 0
Attribute Address.VB_MemberFlags = "400"
If IPAddressHandle <> 0 Then
    If SendMessage(IPAddressHandle, IPM_ISBLANK, 0, ByVal 0&) = 0 Then
        Dim Buffer As Long
        SendMessage IPAddressHandle, IPM_GETADDRESS, 0, ByVal VarPtr(Buffer)
        Address = HiByte(HiWord(Buffer)) & "." & LoByte(HiWord(Buffer)) & "." & HiByte(LoWord(Buffer)) & "." & LoByte(LoWord(Buffer))
    End If
End If
End Property

Public Property Let Address(ByVal Value As String)
If IPAddressHandle <> 0 Then
    If Not Value = vbNullString Then
        On Error GoTo Cancel
        Dim AddrValue(1 To 4) As Byte
        Dim Start As Long, Field As Long
        Do
            Field = Field + 1
            Start = InStr(Value, ".")
            If Start = 0 Then Exit Do
            AddrValue(Field) = CByte(Left(Value, Start - 1))
            Value = Mid(Value, Start + 1)
        Loop
        AddrValue(Field) = CByte(Left(Value, Len(Value)))
        On Error GoTo 0
        SendMessage IPAddressHandle, IPM_SETADDRESS, 0, ByVal MakeDWord(MakeWord(AddrValue(4), AddrValue(3)), MakeWord(AddrValue(2), AddrValue(1)))
    Else
        SendMessage IPAddressHandle, IPM_CLEARADDRESS, 0, ByVal 0&
    End If
    RaiseEvent FieldChange(1)
    RaiseEvent FieldChange(2)
    RaiseEvent FieldChange(3)
    RaiseEvent FieldChange(4)
End If
Exit Property
Cancel:
Err.Raise 380
End Property

Public Sub SetFocusToField(ByVal Field As Long)
Attribute SetFocusToField.VB_Description = "Sets the keyboard focus to the specified field in the IP address control. All of the text in that field will be selected."
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
UserControl.SetFocus
If IPAddressHandle <> 0 Then SendMessage IPAddressHandle, IPM_SETFOCUS, Field - 1, ByVal 0&
End Sub

Public Function IsEmptyField(ByVal Field As Long) As Boolean
Attribute IsEmptyField.VB_Description = "Determines whether the specified field is empty or not."
If Field > 4 Or Field < 1 Then Err.Raise Number:=35600, Description:="Field out of bounds"
If IPAddressHandle <> 0 Then IsEmptyField = CBool(SendMessage(IPAddressEditHandle(Field), WM_GETTEXTLENGTH, 0, ByVal 0&) = 0)
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_MOUSEACTIVATE
        Static InProc As Boolean
        If IPAddressTopDesignMode = False And GetFocus() <> IPAddressHandle And (GetFocus() <> IPAddressEditHandle(1) Or IPAddressEditHandle(1) = 0) And (GetFocus() <> IPAddressEditHandle(2) Or IPAddressEditHandle(2) = 0) And (GetFocus() <> IPAddressEditHandle(3) Or IPAddressEditHandle(3) = 0) And (GetFocus() <> IPAddressEditHandle(4) Or IPAddressEditHandle(4) = 0) Then
            If InProc = True Or LoWord(lParam) = HTBORDER Then WindowProcControl = MA_NOACTIVATEANDEAT: Exit Function
            Select Case HiWord(lParam)
                Case WM_LBUTTONDOWN
                    On Error Resume Next
                    With UserControl
                    If .Extender.CausesValidation = True Then
                        InProc = True
                        Call ComCtlsTopParentValidateControls(Me)
                        InProc = False
                        If Err.Number = 380 Then
                            WindowProcControl = MA_NOACTIVATEANDEAT
                        Else
                            SetFocusAPI .hWnd
                            WindowProcControl = MA_NOACTIVATE
                        End If
                    Else
                        SetFocusAPI .hWnd
                        WindowProcControl = MA_NOACTIVATE
                    End If
                    End With
                    On Error GoTo 0
                    Exit Function
            End Select
        End If
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
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
    Case WM_COMMAND
        Const EN_CHANGE As Long = &H300
        If HiWord(wParam) = EN_CHANGE Then RaiseEvent Change
    Case WM_CTLCOLOREDIT
        WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        SetBkMode wParam, 1
        SetTextColor wParam, WinColor(PropForeColor)
        Exit Function
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
                If (IPAddressMouseOver(0) = False And PropMouseTrack = True) Or (IPAddressMouseOver(5) = False And PropMouseTrack = True) Then
                    If IPAddressMouseOver(0) = False And PropMouseTrack = True Then IPAddressMouseOver(0) = True
                    If IPAddressMouseOver(5) = False And PropMouseTrack = True Then
                        IPAddressMouseOver(5) = True
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
        End Select
    Case WM_MOUSELEAVE, WM_NCMOUSEMOVE
        If wMsg = WM_NCMOUSEMOVE And IPAddressMouseOver(5) = False Then Exit Function
        Dim TME As TRACKMOUSEEVENTSTRUCT
        With TME
        .cbSize = LenB(TME)
        .hWndTrack = hWnd
        .dwFlags = TME_LEAVE Or TME_NONCLIENT
        End With
        TrackMouseEvent TME
    Case WM_NCMOUSELEAVE
        IPAddressMouseOver(0) = False
        If IPAddressMouseOver(5) = True Then
            Dim Pos As Long, hWndFromPoint As Long
            Pos = GetMessagePos()
            hWndFromPoint = WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos))
            If (hWndFromPoint <> IPAddressEditHandle(1) Or IPAddressEditHandle(1) = 0) And (hWndFromPoint <> IPAddressEditHandle(2) Or IPAddressEditHandle(2) = 0) And (hWndFromPoint <> IPAddressEditHandle(3) Or IPAddressEditHandle(3) = 0) And (hWndFromPoint <> IPAddressEditHandle(4) Or IPAddressEditHandle(4) = 0) Then
                IPAddressMouseOver(5) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Index As Long
If hWnd = IPAddressEditHandle(1) Then
    Index = 1
ElseIf hWnd = IPAddressEditHandle(2) Then
    Index = 2
ElseIf hWnd = IPAddressEditHandle(3) Then
    Index = 3
ElseIf hWnd = IPAddressEditHandle(4) Then
    Index = 4
End If
Select Case wMsg
    Case WM_SETFOCUS
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            IPAddressCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If IPAddressCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(IPAddressCharCodeCache And &HFFFF&)
            IPAddressCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        If (wParam And &HFFFF&) <> 0 And KeyChar = 0 Then
            Exit Function
        Else
            wParam = CIntToUInt(KeyChar)
        End If
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then WindowProcEdit = 1 Else SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_PASTE
        If ComCtlsSupportLevel() <= 1 Then
            Dim Text As String
            Text = GetClipboardText()
            If Not Text = vbNullString Then
                Dim i As Long, InvalidText As Boolean
                For i = 1 To Len(Text)
                    If InStr("0123456789", Mid$(Text, i, 1)) = 0 Then
                        InvalidText = True
                        Exit For
                    End If
                Next i
                If InvalidText = True Then
                    VBA.Interaction.Beep
                    Exit Function
                End If
            End If
        End If
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim P As POINTAPI
        P.X = Get_X_lParam(lParam)
        P.Y = Get_Y_lParam(lParam)
        If IPAddressHandle <> 0 Then MapWindowPoints hWnd, IPAddressHandle, P, 1
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(P.X, vbPixels, vbTwips)
        Y = UserControl.ScaleY(P.Y, vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                If Index > 0 Then
                    If (IPAddressMouseOver(Index) = False And PropMouseTrack = True) Or (IPAddressMouseOver(5) = False And PropMouseTrack = True) Then
                        If IPAddressMouseOver(Index) = False And PropMouseTrack = True Then IPAddressMouseOver(Index) = True
                        If IPAddressMouseOver(5) = False And PropMouseTrack = True Then
                            IPAddressMouseOver(5) = True
                            RaiseEvent MouseEnter
                        End If
                        Call ComCtlsRequestMouseLeave(hWnd)
                    End If
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
        End Select
    Case WM_MOUSELEAVE
        If Index > 0 Then
            IPAddressMouseOver(Index) = False
            If IPAddressMouseOver(5) = True Then
                Dim Pos As Long, hWndFromPoint As Long
                Pos = GetMessagePos()
                hWndFromPoint = WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos))
                If (hWndFromPoint <> IPAddressHandle Or IPAddressHandle = 0) And (hWndFromPoint <> IPAddressEditHandle(1) Or IPAddressEditHandle(1) = 0) And (hWndFromPoint <> IPAddressEditHandle(2) Or IPAddressEditHandle(2) = 0) And (hWndFromPoint <> IPAddressEditHandle(3) Or IPAddressEditHandle(3) = 0) And (hWndFromPoint <> IPAddressEditHandle(4) Or IPAddressEditHandle(4) = 0) Then
                    IPAddressMouseOver(5) = False
                    RaiseEvent MouseLeave
                End If
            End If
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = IPAddressHandle Then
            Select Case NM.Code
                Case IPN_FIELDCHANGED
                    Dim NMIPA As NMIPADDRESS
                    CopyMemory NMIPA, ByVal lParam, LenB(NMIPA)
                    RaiseEvent FieldChange(NMIPA.iField + 1)
            End Select
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS Then SetFocusAPI IPAddressHandle
End Function
