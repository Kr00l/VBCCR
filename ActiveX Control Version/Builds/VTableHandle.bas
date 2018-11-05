Attribute VB_Name = "VTableHandle"
Option Explicit

' Required:

' OLEGuids.tlb (in IDE only)

#If False Then
Private VTableInterfaceControl, VTableInterfaceInPlaceActiveObject, VTableInterfacePerPropertyBrowsing
#End If
Public Enum VTableInterfaceConstants
VTableInterfaceControl = 1
VTableInterfaceInPlaceActiveObject = 2
VTableInterfacePerPropertyBrowsing = 3
End Enum
Private Type VTableIPAODataStruct
VTable As Long
RefCount As Long
OriginalIOleIPAO As OLEGuids.IOleInPlaceActiveObject
IOleIPAO As OLEGuids.IOleInPlaceActiveObjectVB
End Type
Private Enum VTableIndexControlConstants
' Ignore : ControlQueryInterface = 1
' Ignore : ControlAddRef = 2
' Ignore : ControlRelease = 3
VTableIndexControlGetControlInfo = 4
VTableIndexControlOnMnemonic = 5
' Ignore : ControlOnAmbientPropertyChange = 6
' Ignore : ControlFreezeEvents = 7
End Enum
Private Enum VTableIndexPPBConstants
' Ignore : PPBQueryInterface = 1
' Ignore : PPBAddRef = 2
' Ignore : PPBRelease = 3
VTableIndexPPBGetDisplayString = 4
' Ignore : PPBMapPropertyToPage = 5
VTAbleIndexPPBGetPredefinedStrings = 6
VTAbleIndexPPBGetPredefinedValue = 7
End Enum
Private Type VTableIEnumVARIANTDataStruct
VTable As Long
RefCount As Long
Enumerable As Object
Index As Long
Count As Long
End Type
Public Const CTRLINFO_EATS_RETURN As Long = 1
Public Const CTRLINFO_EATS_ESCAPE As Long = 2
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadID As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByVal lpString As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal lpvInstance As Long, ByVal oVft As Long, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByRef pvargResult As Variant) As Long
Private Declare Function VariantCopyToPtr Lib "oleaut32" Alias "VariantCopy" (ByVal pvargDest As Long, ByRef pvargSrc As Variant) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Const CC_STDCALL As Long = 4
Private Const GA_ROOT As Long = 2
Private Const GWL_HWNDPARENT As Long = (-8)
Private Const E_OUTOFMEMORY As Long = &H8007000E
Private Const E_INVALIDARG As Long = &H80070057
Private Const E_NOTIMPL As Long = &H80004001
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0
Private VTableIPAO(0 To 9) As Long, VTableIPAOData As VTableIPAODataStruct
Private VTableSubclassControl As VTableSubclass
Private VTableSubclassPPB As VTableSubclass, StringsOutArray() As String, CookiesOutArray() As Long
Private VTableIEnumVARIANT(0 To 6) As Long

Public Sub SetVTableSubclass(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants)
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableSubclassSupported(This, VTableInterfaceInPlaceActiveObject) = True Then VTableIPAOData.RefCount = VTableIPAOData.RefCount + 1
    Case VTableInterfaceControl
        If VTableSubclassSupported(This, VTableInterfaceControl) = True Then Call ReplaceIOleControl(This)
    Case VTableInterfacePerPropertyBrowsing
        If VTableSubclassSupported(This, VTableInterfacePerPropertyBrowsing) = True Then Call ReplaceIPPB(This)
End Select
End Sub

Public Sub RemoveVTableSubclass(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants)
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableSubclassSupported(This, VTableInterfaceInPlaceActiveObject) = True Then VTableIPAOData.RefCount = VTableIPAOData.RefCount - 1
    Case VTableInterfaceControl
        If VTableSubclassSupported(This, VTableInterfaceControl) = True Then Call RestoreIOleControl(This)
    Case VTableInterfacePerPropertyBrowsing
        If VTableSubclassSupported(This, VTableInterfacePerPropertyBrowsing) = True Then Call RestoreIPPB(This)
End Select
End Sub

Public Sub RemoveAllVTableSubclass(ByVal OLEInterface As VTableInterfaceConstants)
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        VTableIPAOData.RefCount = 0
        If Not VTableIPAOData.OriginalIOleIPAO Is Nothing Then Call ActivateIPAO(VTableIPAOData.OriginalIOleIPAO)
    Case VTableInterfaceControl
        Set VTableSubclassControl = Nothing
    Case VTableInterfacePerPropertyBrowsing
        Set VTableSubclassPPB = Nothing
End Select
End Sub

Private Function VTableSubclassSupported(ByRef This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
On Error GoTo CATCH_EXCEPTION
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        Dim ShadowIOleIPAO As OLEGuids.IOleInPlaceActiveObject
        Dim ShadowIOleInPlaceActiveObjectVB As OLEGuids.IOleInPlaceActiveObjectVB
        Set ShadowIOleIPAO = This
        Set ShadowIOleInPlaceActiveObjectVB = This
        VTableSubclassSupported = Not CBool(ShadowIOleIPAO Is Nothing Or ShadowIOleInPlaceActiveObjectVB Is Nothing)
    Case VTableInterfaceControl
        Dim ShadowIOleControl As OLEGuids.IOleControl
        Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
        Set ShadowIOleControl = This
        Set ShadowIOleControlVB = This
        VTableSubclassSupported = Not CBool(ShadowIOleControl Is Nothing Or ShadowIOleControlVB Is Nothing)
    Case VTableInterfacePerPropertyBrowsing
        Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
        Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
        Set ShadowIPPB = This
        Set ShadowIPerPropertyBrowsingVB = This
        VTableSubclassSupported = Not CBool(ShadowIPPB Is Nothing Or ShadowIPerPropertyBrowsingVB Is Nothing)
End Select
CATCH_EXCEPTION:
End Function

Public Function VTableCall(ByVal RetType As VbVarType, ByVal InterfacePointer As Long, ByVal Entry As Long, ParamArray ArgList() As Variant) As Variant
Debug.Assert Not (Entry < 1 Or InterfacePointer = 0)
Dim VarArgList As Variant, HResult As Long
VarArgList = ArgList
If UBound(VarArgList) > -1 Then
    Dim i As Long, ArrVarType() As Integer, ArrVarPtr() As Long
    ReDim ArrVarType(LBound(VarArgList) To UBound(VarArgList)) As Integer
    ReDim ArrVarPtr(LBound(VarArgList) To UBound(VarArgList)) As Long
    For i = LBound(VarArgList) To UBound(VarArgList)
        ArrVarType(i) = VarType(VarArgList(i))
        ArrVarPtr(i) = VarPtr(VarArgList(i))
    Next i
    HResult = DispCallFunc(InterfacePointer, (Entry - 1) * 4, CC_STDCALL, RetType, i, VarPtr(ArrVarType(0)), VarPtr(ArrVarPtr(0)), VTableCall)
Else
    HResult = DispCallFunc(InterfacePointer, (Entry - 1) * 4, CC_STDCALL, RetType, 0, 0, 0, VTableCall)
End If
SetLastError HResult ' S_OK will clear the last error code, if any.
End Function

Public Function VTableInterfaceSupported(ByVal This As OLEGuids.IUnknownUnrestricted, ByVal IIDString As String) As Boolean
Debug.Assert Not (This Is Nothing)
Dim HResult As Long, IID As OLEGuids.OLECLSID, ObjectPointer As Long
CLSIDFromString StrPtr(IIDString), IID
HResult = This.QueryInterface(VarPtr(IID), ObjectPointer)
If ObjectPointer <> 0 Then
    Dim IUnk As OLEGuids.IUnknownUnrestricted
    CopyMemory IUnk, ObjectPointer, 4
    IUnk.Release
    CopyMemory IUnk, 0&, 4
End If
VTableInterfaceSupported = CBool(HResult = S_OK)
End Function

Public Sub ActivateIPAO(ByVal This As Object)
On Error GoTo CATCH_EXCEPTION
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
Dim PropOleInPlaceActiveObject As OLEGuids.IOleInPlaceActiveObject
Dim PosRect As OLEGuids.OLERECT
Dim ClipRect As OLEGuids.OLERECT
Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
Set PropOleObject = This
If VTableIPAOData.RefCount > 0 Then
    With VTableIPAOData
    .VTable = GetVTableIPAO()
    Set .OriginalIOleIPAO = This
    Set .IOleIPAO = This
    End With
    CopyMemory ByVal VarPtr(PropOleInPlaceActiveObject), VarPtr(VTableIPAOData), 4
    PropOleInPlaceActiveObject.AddRef
Else
    Set PropOleInPlaceActiveObject = This
End If
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject PropOleInPlaceActiveObject, vbNullString
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject PropOleInPlaceActiveObject, vbNullString
CATCH_EXCEPTION:
End Sub

Public Sub DeActivateIPAO()
On Error GoTo CATCH_EXCEPTION
If VTableIPAOData.OriginalIOleIPAO Is Nothing Then Exit Sub
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
Dim PropOleInPlaceActiveObject As OLEGuids.IOleInPlaceActiveObject
Dim PosRect As OLEGuids.OLERECT
Dim ClipRect As OLEGuids.OLERECT
Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
Set PropOleObject = VTableIPAOData.OriginalIOleIPAO
Set PropOleInPlaceActiveObject = VTableIPAOData.OriginalIOleIPAO
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject PropOleInPlaceActiveObject, vbNullString
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject PropOleInPlaceActiveObject, vbNullString
Set VTableIPAOData.OriginalIOleIPAO = Nothing
Set VTableIPAOData.IOleIPAO = Nothing
CATCH_EXCEPTION:
End Sub

Private Function GetVTableIPAO() As Long
If VTableIPAO(0) = 0 Then
    VTableIPAO(0) = ProcPtr(AddressOf IOleIPAO_QueryInterface)
    VTableIPAO(1) = ProcPtr(AddressOf IOleIPAO_AddRef)
    VTableIPAO(2) = ProcPtr(AddressOf IOleIPAO_Release)
    VTableIPAO(3) = ProcPtr(AddressOf IOleIPAO_GetWindow)
    VTableIPAO(4) = ProcPtr(AddressOf IOleIPAO_ContextSensitiveHelp)
    VTableIPAO(5) = ProcPtr(AddressOf IOleIPAO_TranslateAccelerator)
    VTableIPAO(6) = ProcPtr(AddressOf IOleIPAO_OnFrameWindowActivate)
    VTableIPAO(7) = ProcPtr(AddressOf IOleIPAO_OnDocWindowActivate)
    VTableIPAO(8) = ProcPtr(AddressOf IOleIPAO_ResizeBorder)
    VTableIPAO(9) = ProcPtr(AddressOf IOleIPAO_EnableModeless)
End If
GetVTableIPAO = VarPtr(VTableIPAO(0))
End Function

Private Function IOleIPAO_QueryInterface(ByRef This As VTableIPAODataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As Long) As Long
If VarPtr(pvObj) = 0 Then
    IOleIPAO_QueryInterface = E_POINTER
    Exit Function
End If
' IID_IOleInPlaceActiveObject = {00000117-0000-0000-C000-000000000046}
If IID.Data1 = &H117 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
    If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
    And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
        pvObj = VarPtr(This)
        IOleIPAO_AddRef This
        IOleIPAO_QueryInterface = S_OK
    Else
        IOleIPAO_QueryInterface = This.OriginalIOleIPAO.QueryInterface(VarPtr(IID), pvObj)
    End If
Else
    IOleIPAO_QueryInterface = This.OriginalIOleIPAO.QueryInterface(VarPtr(IID), pvObj)
End If
End Function

Private Function IOleIPAO_AddRef(ByRef This As VTableIPAODataStruct) As Long
IOleIPAO_AddRef = This.OriginalIOleIPAO.AddRef
End Function

Private Function IOleIPAO_Release(ByRef This As VTableIPAODataStruct) As Long
IOleIPAO_Release = This.OriginalIOleIPAO.Release
End Function

Private Function IOleIPAO_GetWindow(ByRef This As VTableIPAODataStruct, ByRef hWnd As Long) As Long
IOleIPAO_GetWindow = This.OriginalIOleIPAO.GetWindow(hWnd)
End Function

Private Function IOleIPAO_ContextSensitiveHelp(ByRef This As VTableIPAODataStruct, ByVal EnterMode As Long) As Long
IOleIPAO_ContextSensitiveHelp = This.OriginalIOleIPAO.ContextSensitiveHelp(EnterMode)
End Function

Private Function IOleIPAO_TranslateAccelerator(ByRef This As VTableIPAODataStruct, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
If VarPtr(Msg) = 0 Then
    IOleIPAO_TranslateAccelerator = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim Handled As Boolean
IOleIPAO_TranslateAccelerator = S_OK
This.IOleIPAO.TranslateAccelerator Handled, IOleIPAO_TranslateAccelerator, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
If Handled = False Then IOleIPAO_TranslateAccelerator = This.OriginalIOleIPAO.TranslateAccelerator(VarPtr(Msg))
Exit Function
CATCH_EXCEPTION:
IOleIPAO_TranslateAccelerator = This.OriginalIOleIPAO.TranslateAccelerator(VarPtr(Msg))
End Function

Private Function IOleIPAO_OnFrameWindowActivate(ByRef This As VTableIPAODataStruct, ByVal Activate As Long) As Long
IOleIPAO_OnFrameWindowActivate = This.OriginalIOleIPAO.OnFrameWindowActivate(Activate)
End Function

Private Function IOleIPAO_OnDocWindowActivate(ByRef This As VTableIPAODataStruct, ByVal Activate As Long) As Long
IOleIPAO_OnDocWindowActivate = This.OriginalIOleIPAO.OnDocWindowActivate(Activate)
End Function

Private Function IOleIPAO_ResizeBorder(ByRef This As VTableIPAODataStruct, ByRef RC As OLEGuids.OLERECT, ByVal UIWindow As OLEGuids.IOleInPlaceUIWindow, ByVal FrameWindow As Long) As Long
IOleIPAO_ResizeBorder = This.OriginalIOleIPAO.ResizeBorder(VarPtr(RC), UIWindow, FrameWindow)
End Function

Private Function IOleIPAO_EnableModeless(ByRef This As VTableIPAODataStruct, ByVal Enable As Long) As Long
IOleIPAO_EnableModeless = This.OriginalIOleIPAO.EnableModeless(Enable)
End Function

Private Sub ReplaceIOleControl(ByVal This As OLEGuids.IOleControl)
If VTableSubclassControl Is Nothing Then Set VTableSubclassControl = New VTableSubclass
If VTableSubclassControl.RefCount = 0 Then
    Dim hMain As Long, Handled As Boolean
    hMain = GetHiddenMainWindow()
    If hMain <> 0 Then Handled = CBool(GetProp(hMain, StrPtr("VTableSubclassControlInit")) <> 0)
    If Handled = False Then
        VTableSubclassControl.Subclass ObjPtr(This), VTableIndexControlGetControlInfo, VTableIndexControlOnMnemonic, _
        AddressOf IOleControl_GetControlInfo, _
        AddressOf IOleControl_OnMnemonic
        If hMain <> 0 Then SetProp hMain, StrPtr("VTableSubclassControlInit"), 1
    End If
End If
VTableSubclassControl.AddRef
End Sub

Private Sub RestoreIOleControl(ByVal This As OLEGuids.IOleControl)
If Not VTableSubclassControl Is Nothing Then
    VTableSubclassControl.Release
    If VTableSubclassControl.RefCount = 0 Then
        Dim hMain As Long
        hMain = GetHiddenMainWindow()
        If hMain <> 0 Then RemoveProp hMain, StrPtr("VTableSubclassControlInit")
        VTableSubclassControl.UnSubclass
    End If
End If
End Sub

Public Sub OnControlInfoChanged(ByVal This As Object, Optional ByVal OnFocus As Boolean)
Dim PropOleObject As OLEGuids.IOleObject
Dim PropClientSite As OLEGuids.IOleClientSite
Dim PropUnknown As IUnknown
Dim PropControlSite As OLEGuids.IOleControlSite
On Error Resume Next
Set PropOleObject = This
Set PropClientSite = PropOleObject.GetClientSite
Set PropUnknown = PropClientSite
Set PropControlSite = PropUnknown
PropControlSite.OnControlInfoChanged
If OnFocus = True Then PropControlSite.OnFocus 1
End Sub

Private Function IOleControl_GetControlInfo(ByVal This As Object, ByRef CI As OLEGuids.OLECONTROLINFO) As Long
If VarPtr(CI) = 0 Then
    IOleControl_GetControlInfo = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIOleControlVB As OLEGuids.IOleControlVB, Handled As Boolean
Set ShadowIOleControlVB = This
CI.cb = LenB(CI)
ShadowIOleControlVB.GetControlInfo Handled, CI.cAccel, CI.hAccel, CI.dwFlags
If Handled = False Then
    IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
Else
    If CI.cAccel > 0 And CI.hAccel = 0 Then
        IOleControl_GetControlInfo = E_OUTOFMEMORY
    Else
        IOleControl_GetControlInfo = S_OK
    End If
End If
Exit Function
CATCH_EXCEPTION:
IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
End Function

Private Function IOleControl_OnMnemonic(ByVal This As Object, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
If VarPtr(Msg) = 0 Then
    IOleControl_OnMnemonic = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIOleControlVB As OLEGuids.IOleControlVB, Handled As Boolean
Set ShadowIOleControlVB = This
ShadowIOleControlVB.OnMnemonic Handled, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
If Handled = False Then
    IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
Else
    IOleControl_OnMnemonic = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
End Function

Private Function Original_IOleControl_GetControlInfo(ByVal This As OLEGuids.IOleControl, ByRef CI As OLEGuids.OLECONTROLINFO) As Long
VTableSubclassControl.SubclassEntry(VTableIndexControlGetControlInfo) = False
Original_IOleControl_GetControlInfo = This.GetControlInfo(CI)
VTableSubclassControl.SubclassEntry(VTableIndexControlGetControlInfo) = True
End Function

Private Function Original_IOleControl_OnMnemonic(ByVal This As OLEGuids.IOleControl, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
VTableSubclassControl.SubclassEntry(VTableIndexControlOnMnemonic) = False
Original_IOleControl_OnMnemonic = This.OnMnemonic(Msg)
VTableSubclassControl.SubclassEntry(VTableIndexControlOnMnemonic) = True
End Function

Private Sub ReplaceIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
If VTableSubclassPPB Is Nothing Then Set VTableSubclassPPB = New VTableSubclass
If VTableSubclassPPB.RefCount = 0 Then
    Dim hMain As Long, Handled As Boolean
    hMain = GetHiddenMainWindow()
    If hMain <> 0 Then Handled = CBool(GetProp(hMain, StrPtr("VTableSubclassPPBInit")) <> 0)
    If Handled = False Then
        VTableSubclassPPB.Subclass ObjPtr(This), VTableIndexPPBGetDisplayString, VTAbleIndexPPBGetPredefinedValue, _
        AddressOf IPPB_GetDisplayString, 0, _
        AddressOf IPPB_GetPredefinedStrings, _
        AddressOf IPPB_GetPredefinedValue
        If hMain <> 0 Then SetProp hMain, StrPtr("VTableSubclassPPBInit"), 1
    End If
End If
VTableSubclassPPB.AddRef
End Sub

Private Sub RestoreIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
If Not VTableSubclassPPB Is Nothing Then
    VTableSubclassPPB.Release
    If VTableSubclassPPB.RefCount = 0 Then
        Dim hMain As Long
        hMain = GetHiddenMainWindow()
        If hMain <> 0 Then RemoveProp hMain, StrPtr("VTableSubclassPPBInit")
        VTableSubclassPPB.UnSubclass
    End If
End If
End Sub

Public Function GetDispID(ByVal This As Object, ByRef MethodName As String) As Long
Dim IDispatch As OLEGuids.IDispatch, IID_NULL As OLEGuids.OLECLSID
Set IDispatch = This
IDispatch.GetIDsOfNames IID_NULL, MethodName, 1, 0, GetDispID
End Function

Private Function IPPB_GetDisplayString(ByVal This As Object, ByVal DispID As Long, ByVal lpDisplayName As Long) As Long
If lpDisplayName = 0 Then
    IPPB_GetDisplayString = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB, Handled As Boolean, DisplayName As String
Set ShadowIPerPropertyBrowsingVB = This
ShadowIPerPropertyBrowsingVB.GetDisplayString Handled, DispID, DisplayName
If Handled = False Then
    IPPB_GetDisplayString = E_NOTIMPL
Else
    Dim lpString As Long
    lpString = SysAllocString(StrPtr(DisplayName))
    CopyMemory ByVal lpDisplayName, lpString, 4
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetDisplayString = E_NOTIMPL
End Function

Private Function IPPB_GetPredefinedStrings(ByVal This As Object, ByVal DispID As Long, ByRef pCaStringsOut As OLEGuids.OLECALPOLESTR, ByRef pCaCookiesOut As OLEGuids.OLECADWORD) As Long
If VarPtr(pCaStringsOut) = 0 Or VarPtr(pCaCookiesOut) = 0 Then
    IPPB_GetPredefinedStrings = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB, Handled As Boolean
ReDim StringsOutArray(0) As String
ReDim CookiesOutArray(0) As Long
Set ShadowIPerPropertyBrowsingVB = This
ShadowIPerPropertyBrowsingVB.GetPredefinedStrings Handled, DispID, StringsOutArray(), CookiesOutArray()
If Handled = False Or UBound(StringsOutArray()) = 0 Then
    IPPB_GetPredefinedStrings = E_NOTIMPL
Else
    Dim cElems As Long, pElems As Long, nElemCount As Long
    Dim lpString As Long
    cElems = UBound(StringsOutArray())
    If Not UBound(CookiesOutArray()) = cElems Then ReDim Preserve CookiesOutArray(cElems) As Long
    pElems = CoTaskMemAlloc(cElems * 4)
    pCaStringsOut.cElems = cElems
    pCaStringsOut.pElems = pElems
    For nElemCount = 0 To cElems - 1
        lpString = CoTaskMemAlloc(Len(StringsOutArray(nElemCount)) + 1)
        CopyMemory ByVal lpString, StrPtr(StringsOutArray(nElemCount)), 4
        CopyMemory ByVal UnsignedAdd(pElems, nElemCount * 4), ByVal lpString, 4
    Next nElemCount
    pElems = CoTaskMemAlloc(cElems * 4)
    pCaCookiesOut.cElems = cElems
    pCaCookiesOut.pElems = pElems
    For nElemCount = 0 To cElems - 1
        CopyMemory ByVal UnsignedAdd(pElems, nElemCount * 4), CookiesOutArray(nElemCount), 4
    Next nElemCount
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetPredefinedStrings = E_NOTIMPL
End Function

Private Function IPPB_GetPredefinedValue(ByVal This As Object, ByVal DispID As Long, ByVal dwCookie As Long, ByRef pVarOut As Variant) As Long
If VarPtr(pVarOut) = 0 Then
    IPPB_GetPredefinedValue = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB, Handled As Boolean
Set ShadowIPerPropertyBrowsingVB = This
ShadowIPerPropertyBrowsingVB.GetPredefinedValue Handled, DispID, dwCookie, pVarOut
If Handled = False Then IPPB_GetPredefinedValue = E_NOTIMPL
Exit Function
CATCH_EXCEPTION:
IPPB_GetPredefinedValue = E_NOTIMPL
End Function

Public Function GetNewEnum(ByVal This As Object, ByVal Upper As Long, ByVal Lower As Long) As IEnumVARIANT
Dim VTableIEnumVARIANTData As VTableIEnumVARIANTDataStruct
With VTableIEnumVARIANTData
.VTable = GetVTableIEnumVARIANT()
.RefCount = 1
Set .Enumerable = This
.Index = Lower
.Count = Upper
Dim hMem As Long
hMem = CoTaskMemAlloc(LenB(VTableIEnumVARIANTData))
If hMem <> 0 Then
    CopyMemory ByVal hMem, VTableIEnumVARIANTData, LenB(VTableIEnumVARIANTData)
    CopyMemory ByVal VarPtr(GetNewEnum), hMem, 4
    CopyMemory ByVal VarPtr(.Enumerable), 0&, 4
End If
End With
End Function

Private Function GetVTableIEnumVARIANT() As Long
If VTableIEnumVARIANT(0) = 0 Then
    VTableIEnumVARIANT(0) = ProcPtr(AddressOf IEnumVARIANT_QueryInterface)
    VTableIEnumVARIANT(1) = ProcPtr(AddressOf IEnumVARIANT_AddRef)
    VTableIEnumVARIANT(2) = ProcPtr(AddressOf IEnumVARIANT_Release)
    VTableIEnumVARIANT(3) = ProcPtr(AddressOf IEnumVARIANT_Next)
    VTableIEnumVARIANT(4) = ProcPtr(AddressOf IEnumVARIANT_Skip)
    VTableIEnumVARIANT(5) = ProcPtr(AddressOf IEnumVARIANT_Reset)
    VTableIEnumVARIANT(6) = ProcPtr(AddressOf IEnumVARIANT_Clone)
End If
GetVTableIEnumVARIANT = VarPtr(VTableIEnumVARIANT(0))
End Function

Private Function IEnumVARIANT_QueryInterface(ByRef This As VTableIEnumVARIANTDataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As Long) As Long
If VarPtr(pvObj) = 0 Then
    IEnumVARIANT_QueryInterface = E_POINTER
    Exit Function
End If
' IID_IEnumVARIANT = {00020404-0000-0000-C000-000000000046}
If IID.Data1 = &H20404 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
    If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
    And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
        pvObj = VarPtr(This)
        IEnumVARIANT_AddRef This
        IEnumVARIANT_QueryInterface = S_OK
    Else
        IEnumVARIANT_QueryInterface = E_NOINTERFACE
    End If
Else
    IEnumVARIANT_QueryInterface = E_NOINTERFACE
End If
End Function

Private Function IEnumVARIANT_AddRef(ByRef This As VTableIEnumVARIANTDataStruct) As Long
This.RefCount = This.RefCount + 1
IEnumVARIANT_AddRef = This.RefCount
End Function

Private Function IEnumVARIANT_Release(ByRef This As VTableIEnumVARIANTDataStruct) As Long
This.RefCount = This.RefCount - 1
IEnumVARIANT_Release = This.RefCount
If IEnumVARIANT_Release = 0 Then
    Set This.Enumerable = Nothing
    CoTaskMemFree VarPtr(This)
End If
End Function

Private Function IEnumVARIANT_Next(ByRef This As VTableIEnumVARIANTDataStruct, ByVal VntCount As Long, ByVal VntArrPtr As Long, ByRef pcvFetched As Long) As Long
If VntArrPtr = 0 Then
    IEnumVARIANT_Next = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Const VARIANT_CB As Long = 16
Dim Fetched As Long
With This
Do Until .Index > .Count
    VariantCopyToPtr VntArrPtr, .Enumerable(.Index)
    .Index = .Index + 1
    Fetched = Fetched + 1
    If Fetched = VntCount Then Exit Do
    VntArrPtr = UnsignedAdd(VntArrPtr, VARIANT_CB)
Loop
End With
If Fetched = VntCount Then
    IEnumVARIANT_Next = S_OK
Else
    IEnumVARIANT_Next = S_FALSE
End If
If VarPtr(pcvFetched) <> 0 Then pcvFetched = Fetched
Exit Function
CATCH_EXCEPTION:
If VarPtr(pcvFetched) <> 0 Then pcvFetched = 0
IEnumVARIANT_Next = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Skip(ByRef This As VTableIEnumVARIANTDataStruct, ByVal VntCount As Long) As Long
IEnumVARIANT_Skip = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Reset(ByRef This As VTableIEnumVARIANTDataStruct) As Long
IEnumVARIANT_Reset = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Clone(ByRef This As VTableIEnumVARIANTDataStruct, ByRef ppEnum As IEnumVARIANT) As Long
IEnumVARIANT_Clone = E_NOTIMPL
End Function

Private Function GetHiddenMainWindow() As Long
EnumThreadWindows App.ThreadID, AddressOf EnumThreadWndProc, VarPtr(GetHiddenMainWindow)
End Function

Private Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lpResult As Long) As Long
Dim ClassName As String
EnumThreadWndProc = 1
If GetWindowLong(hWnd, GWL_HWNDPARENT) = 0 Then
    ClassName = GetWindowClassName(hWnd)
    If InStr(ClassName, "Thunder") = 1 Then
        If InStr(ClassName, "Main") = (Len(ClassName) - 3) Then
            CopyMemory ByVal lpResult, hWnd, 4
            EnumThreadWndProc = 0
        End If
    End If
End If
End Function
