Attribute VB_Name = "VTableHandle"
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

' Required:

' OLEGuids.tlb (in IDE only)

#If False Then
Private VTableInterfaceInPlaceActiveObject, VTableInterfaceControl, VTableInterfacePerPropertyBrowsing, VTableInterfaceInPlaceObjectWindowless
#End If
Public Enum VTableInterfaceConstants
VTableInterfaceInPlaceActiveObject = 1
VTableInterfaceControl = 2
VTableInterfacePerPropertyBrowsing = 3
VTableInterfaceInPlaceObjectWindowless = 4
End Enum
Private Type VTableIPAODataStruct
VTable As LongPtr
RefCount As Long
OriginalIOleIPAO As OLEGuids.IOleInPlaceActiveObject
IOleIPAO As OLEGuids.IOleInPlaceActiveObjectVB
End Type
Private Type VTableEnumVARIANTDataStruct
VTable As LongPtr
RefCount As Long
Enumerable As Object
Index As Long
Count As Long
End Type
Public Const CTRLINFO_EATS_RETURN As Long = 1
Public Const CTRLINFO_EATS_ESCAPE As Long = 2
#If VBA7 Then
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As LongPtr
Private Declare PtrSafe Function SysAllocString Lib "oleaut32" (ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal lpvInstance As LongPtr, ByVal oVft As LongPtr, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As LongPtr, ByVal prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (ByRef pvargDest As Any, ByRef pvargSrc As Any) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As LongPtr, ByRef pCLSID As Any) As Long
#Else
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByVal lpString As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal lpvInstance As Long, ByVal oVft As Long, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByRef pvargResult As Variant) As Long
Private Declare Function VariantCopy Lib "oleaut32" (ByRef pvargDest As Any, ByRef pvargSrc As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
#End If
Private Const CC_STDCALL As Long = 4
Private Const E_OUTOFMEMORY As Long = &H8007000E
Private Const E_INVALIDARG As Long = &H80070057
Private Const E_NOTIMPL As Long = &H80004001
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0
Private VTableIPAO(0 To 9) As LongPtr, VTableIPAOData As VTableIPAODataStruct
Private VTableControl(0 To 6) As LongPtr, OriginalVTableControl As LongPtr
Private VTablePPB(0 To 6) As LongPtr, OriginalVTablePPB As LongPtr, StringsOutArray() As String, CookiesOutArray() As Long
Private VTableInPlaceObjectWindowless(0 To 10) As LongPtr, OriginalVTableInPlaceObjectWindowless As LongPtr
Private VTableEnumVARIANT(0 To 6) As LongPtr

Public Function SetVTableHandling(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableHandlingSupported(This, VTableInterfaceInPlaceActiveObject) = True Then
            VTableIPAOData.RefCount = VTableIPAOData.RefCount + 1
            SetVTableHandling = True
        End If
    Case VTableInterfaceControl
        If VTableHandlingSupported(This, VTableInterfaceControl) = True Then
            Call ReplaceIOleControl(This)
            SetVTableHandling = True
        End If
    Case VTableInterfacePerPropertyBrowsing
        If VTableHandlingSupported(This, VTableInterfacePerPropertyBrowsing) = True Then
            Call ReplaceIPPB(This)
            SetVTableHandling = True
        End If
    Case VTableInterfaceInPlaceObjectWindowless
        If VTableHandlingSupported(This, VTableInterfaceInPlaceObjectWindowless) = True Then
            Call ReplaceInPlaceObjectWindowless(This)
            SetVTableHandling = True
        End If
End Select
End Function

Public Function RemoveVTableHandling(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableHandlingSupported(This, VTableInterfaceInPlaceActiveObject) = True Then
            VTableIPAOData.RefCount = VTableIPAOData.RefCount - 1
            RemoveVTableHandling = True
        End If
    Case VTableInterfaceControl
        If VTableHandlingSupported(This, VTableInterfaceControl) = True Then
            Call RestoreIOleControl(This)
            RemoveVTableHandling = True
        End If
    Case VTableInterfacePerPropertyBrowsing
        If VTableHandlingSupported(This, VTableInterfacePerPropertyBrowsing) = True Then
            Call RestoreIPPB(This)
            RemoveVTableHandling = True
        End If
    Case VTableInterfaceInPlaceObjectWindowless
        If VTableHandlingSupported(This, VTableInterfaceInPlaceObjectWindowless) = True Then
            Call RestoreInPlaceObjectWindowless(This)
            RemoveVTableHandling = True
        End If
End Select
End Function

Private Function VTableHandlingSupported(ByRef This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
On Error GoTo CATCH_EXCEPTION
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        Dim ShadowIOleIPAO As OLEGuids.IOleInPlaceActiveObject
        Dim ShadowIOleInPlaceActiveObjectVB As OLEGuids.IOleInPlaceActiveObjectVB
        Set ShadowIOleIPAO = This
        Set ShadowIOleInPlaceActiveObjectVB = This
        VTableHandlingSupported = Not CBool(ShadowIOleIPAO Is Nothing Or ShadowIOleInPlaceActiveObjectVB Is Nothing)
    Case VTableInterfaceControl
        Dim ShadowIOleControl As OLEGuids.IOleControl
        Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
        Set ShadowIOleControl = This
        Set ShadowIOleControlVB = This
        VTableHandlingSupported = Not CBool(ShadowIOleControl Is Nothing Or ShadowIOleControlVB Is Nothing)
    Case VTableInterfacePerPropertyBrowsing
        Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
        Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
        Set ShadowIPPB = This
        Set ShadowIPerPropertyBrowsingVB = This
        VTableHandlingSupported = Not CBool(ShadowIPPB Is Nothing Or ShadowIPerPropertyBrowsingVB Is Nothing)
    Case VTableInterfaceInPlaceObjectWindowless
        Dim ShadowIOleInPlaceObjectWindowless As OLEGuids.IOleInPlaceObjectWindowless
        Dim ShadowIOleInPlaceObjectWindowlessVB As OLEGuids.IOleInPlaceObjectWindowlessVB
        Set ShadowIOleInPlaceObjectWindowless = This
        Set ShadowIOleInPlaceObjectWindowlessVB = This
        VTableHandlingSupported = Not CBool(ShadowIOleInPlaceObjectWindowless Is Nothing Or ShadowIOleInPlaceObjectWindowlessVB Is Nothing)
End Select
CATCH_EXCEPTION:
End Function

#If VBA7 Then
Public Function VTableCall(ByVal RetType As VbVarType, ByVal InterfacePointer As LongPtr, ByVal Entry As LongPtr, ParamArray ArgList() As Variant) As Variant
#Else
Public Function VTableCall(ByVal RetType As VbVarType, ByVal InterfacePointer As Long, ByVal Entry As Long, ParamArray ArgList() As Variant) As Variant
#End If
Debug.Assert Not (Entry < 1 Or InterfacePointer = NULL_PTR)
Dim VarArgList As Variant, HResult As Long
VarArgList = ArgList
If UBound(VarArgList) > -1 Then
    Dim i As Long, ArrVarType() As Integer, ArrVarPtr() As LongPtr
    ReDim ArrVarType(LBound(VarArgList) To UBound(VarArgList)) As Integer
    ReDim ArrVarPtr(LBound(VarArgList) To UBound(VarArgList)) ' As LongPtr
    For i = LBound(VarArgList) To UBound(VarArgList)
        ArrVarType(i) = VarType(VarArgList(i))
        ArrVarPtr(i) = VarPtr(VarArgList(i))
    Next i
    HResult = DispCallFunc(InterfacePointer, (Entry - 1) * PTR_SIZE, CC_STDCALL, RetType, i, VarPtr(ArrVarType(0)), VarPtr(ArrVarPtr(0)), VTableCall)
Else
    HResult = DispCallFunc(InterfacePointer, (Entry - 1) * PTR_SIZE, CC_STDCALL, RetType, 0, 0, 0, VTableCall)
End If
SetLastError HResult ' S_OK will clear the last error code, if any.
End Function

Public Function VTableInterfaceSupported(ByVal This As OLEGuids.IUnknownUnrestricted, ByVal IIDString As String) As Boolean
Dim HResult As Long, IID As OLEGuids.OLECLSID, ObjectPointer As LongPtr
CLSIDFromString StrPtr(IIDString), IID
HResult = This.QueryInterface(VarPtr(IID), ObjectPointer)
If ObjectPointer <> NULL_PTR Then
    Dim IUnk As OLEGuids.IUnknownUnrestricted
    CopyMemory IUnk, ObjectPointer, PTR_SIZE
    IUnk.Release
    CopyMemory IUnk, NULL_PTR, PTR_SIZE
End If
VTableInterfaceSupported = CBool(HResult = S_OK)
End Function

Public Function GetDispId(ByVal This As Object, ByRef MethodName As String) As Long
Dim PropDispatch As OLEGuids.IDispatchUnrestricted, IID_NULL As OLEGuids.OLECLSID
Set PropDispatch = This
PropDispatch.GetIDsOfNames IID_NULL, StrPtr(MethodName), 1, 0, GetDispId
End Function

#If (TWINBASIC = 0) Then
Public Function CallByDispId(ByVal This As Object, ByVal DispId As Long, ByVal CallType As VbCallType, ParamArray ArgList() As Variant) As Variant
Const DISPID_PROPERTYPUT As Long = -3
Dim PropDispatch As OLEGuids.IDispatchUnrestricted, IID_NULL As OLEGuids.OLECLSID, pDispParams As OLEGuids.OLEDISPPARAMS
Set PropDispatch = This
If UBound(ArgList) > -1 Then
    Dim i As Long, ArgListRev As Variant
    ReDim ArgListRev(LBound(ArgList) To UBound(ArgList)) As Variant
    For i = LBound(ArgList) To UBound(ArgList)
        VariantCopy ArgListRev(i), ArgList(UBound(ArgList) - i)
    Next i
    pDispParams.rgvarg = VarPtr(ArgListRev(0))
    pDispParams.cArgs = UBound(ArgListRev) + 1
End If
Dim PropPutDispId As Long
If (CallType And (VbLet Or VbSet)) <> 0 Then
    PropPutDispId = DISPID_PROPERTYPUT
    pDispParams.rgdispidNamedArgs = VarPtr(PropPutDispId)
    pDispParams.cNamedArgs = 1
End If
PropDispatch.Invoke DispId, IID_NULL, 0, CallType, VarPtr(pDispParams), VarPtr(CallByDispId), NULL_PTR, 0&
End Function
#End If

#If VBA7 Then
Public Function GetWindowFromObject(ByVal This As Object) As LongPtr
#Else
Public Function GetWindowFromObject(ByVal This As Object) As Long
#End If
If This Is Nothing Then Exit Function
If TypeOf This Is OLEGuids.IOleWindow Then
    Dim PropOleWindow As OLEGuids.IOleWindow
    Set PropOleWindow = This
    GetWindowFromObject = PropOleWindow.GetWindow()
End If
End Function

Public Sub SyncObjectRectsToContainer(ByVal This As Object)
On Error GoTo CATCH_EXCEPTION
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceObject As OLEGuids.IOleInPlaceObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PosRect As OLEGuids.OLERECT
Dim ClipRect As OLEGuids.OLERECT
Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
Set PropOleObject = This
Set PropOleInPlaceObject = This
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext Nothing, Nothing, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceObject.SetObjectRects VarPtr(PosRect), VarPtr(ClipRect)
CATCH_EXCEPTION:
End Sub

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
    CopyMemory ByVal VarPtr(PropOleInPlaceActiveObject), VarPtr(VTableIPAOData), PTR_SIZE
    PropOleInPlaceActiveObject.AddRef
Else
    Set PropOleInPlaceActiveObject = This
End If
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject PropOleInPlaceActiveObject, NULL_PTR
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject PropOleInPlaceActiveObject, NULL_PTR
CATCH_EXCEPTION:
End Sub

Public Sub DeActivateIPAO()
On Error GoTo CATCH_EXCEPTION
If VTableIPAOData.OriginalIOleIPAO Is Nothing Then Exit Sub
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
Dim PosRect As OLEGuids.OLERECT
Dim ClipRect As OLEGuids.OLERECT
Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
Set PropOleObject = VTableIPAOData.OriginalIOleIPAO
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject Nothing, NULL_PTR
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject Nothing, NULL_PTR
CATCH_EXCEPTION:
Set VTableIPAOData.OriginalIOleIPAO = Nothing
Set VTableIPAOData.IOleIPAO = Nothing
End Sub

Private Function GetVTableIPAO() As LongPtr
If VTableIPAO(0) = NULL_PTR Then
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

Private Function IOleIPAO_QueryInterface(ByRef This As VTableIPAODataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As LongPtr) As Long
If VarPtr(pvObj) = NULL_PTR Then
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

Private Function IOleIPAO_GetWindow(ByRef This As VTableIPAODataStruct, ByRef hWnd As LongPtr) As Long
IOleIPAO_GetWindow = This.OriginalIOleIPAO.GetWindow(hWnd)
End Function

Private Function IOleIPAO_ContextSensitiveHelp(ByRef This As VTableIPAODataStruct, ByVal EnterMode As Long) As Long
IOleIPAO_ContextSensitiveHelp = This.OriginalIOleIPAO.ContextSensitiveHelp(EnterMode)
End Function

Private Function IOleIPAO_TranslateAccelerator(ByRef This As VTableIPAODataStruct, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
If VarPtr(Msg) = NULL_PTR Then
    IOleIPAO_TranslateAccelerator = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim Handled As Boolean
IOleIPAO_TranslateAccelerator = S_OK
This.IOleIPAO.TranslateAccelerator Handled, IOleIPAO_TranslateAccelerator, Msg.hWnd, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
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
If OriginalVTableControl = NULL_PTR Then CopyMemory OriginalVTableControl, ByVal ObjPtr(This), PTR_SIZE
If OriginalVTableControl <> NULL_PTR Then CopyMemory ByVal ObjPtr(This), ByVal VarPtr(GetVTableControl()), PTR_SIZE
End Sub

Private Sub RestoreIOleControl(ByVal This As OLEGuids.IOleControl)
If OriginalVTableControl <> NULL_PTR Then CopyMemory ByVal ObjPtr(This), OriginalVTableControl, PTR_SIZE
End Sub

Public Sub OnControlInfoChanged(ByVal This As Object, Optional ByVal OnFocus As Boolean)
On Error GoTo CATCH_EXCEPTION
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleControlSite As OLEGuids.IOleControlSite
Set PropOleObject = This
Set PropOleControlSite = PropOleObject.GetClientSite
PropOleControlSite.OnControlInfoChanged
If OnFocus = True Then PropOleControlSite.OnFocus 1
CATCH_EXCEPTION:
End Sub

Private Function GetVTableControl() As LongPtr
If VTableControl(0) = NULL_PTR Then
    If OriginalVTableControl <> NULL_PTR Then CopyMemory VTableControl(0), ByVal OriginalVTableControl, 3 * PTR_SIZE
    VTableControl(3) = ProcPtr(AddressOf IOleControl_GetControlInfo)
    VTableControl(4) = ProcPtr(AddressOf IOleControl_OnMnemonic)
    VTableControl(5) = ProcPtr(AddressOf IOleControl_OnAmbientPropertyChange)
    If OriginalVTableControl <> NULL_PTR Then
        CopyMemory VTableControl(6), ByVal UnsignedAdd(OriginalVTableControl, 6 * PTR_SIZE), PTR_SIZE
    Else
        VTableControl(6) = ProcPtr(AddressOf IOleControl_FreezeEvents)
    End If
End If
GetVTableControl = VarPtr(VTableControl(0))
End Function

Private Function IOleControl_GetControlInfo(ByRef This As LongPtr, ByRef CI As OLEGuids.OLECONTROLINFO) As Long
If VarPtr(CI) = NULL_PTR Then
    IOleControl_GetControlInfo = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIOleControlVB As OLEGuids.IOleControlVB, Handled As Boolean
Set ShadowIOleControlVB = PtrToObj(VarPtr(This))
CI.cb = LenB(CI)
ShadowIOleControlVB.GetControlInfo Handled, CI.cAccel, CI.hAccel, CI.dwFlags
If Handled = False Then
    IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
Else
    If CI.cAccel > 0 And CI.hAccel = NULL_PTR Then
        IOleControl_GetControlInfo = E_OUTOFMEMORY
    Else
        IOleControl_GetControlInfo = S_OK
    End If
End If
Exit Function
CATCH_EXCEPTION:
IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
End Function

Private Function IOleControl_OnMnemonic(ByRef This As LongPtr, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
If VarPtr(Msg) = NULL_PTR Then
    IOleControl_OnMnemonic = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIOleControlVB As OLEGuids.IOleControlVB, Handled As Boolean
Set ShadowIOleControlVB = PtrToObj(VarPtr(This))
ShadowIOleControlVB.OnMnemonic Handled, Msg.hWnd, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
If Handled = False Then
    IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
Else
    IOleControl_OnMnemonic = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
End Function

Private Function IOleControl_OnAmbientPropertyChange(ByRef This As LongPtr, ByVal DispId As Long) As Long
IOleControl_OnAmbientPropertyChange = Original_IOleControl_OnAmbientPropertyChange(This, DispId)
End Function

Private Function IOleControl_FreezeEvents(ByRef This As LongPtr, ByVal bFreeze As Long) As Long
IOleControl_FreezeEvents = Original_IOleControl_FreezeEvents(This, bFreeze)
End Function

Private Function Original_IOleControl_GetControlInfo(ByRef This As LongPtr, ByRef CI As OLEGuids.OLECONTROLINFO) As Long
If OriginalVTableControl <> NULL_PTR Then
    Dim ShadowIOleControl As OLEGuids.IOleControl
    This = OriginalVTableControl
    CopyMemory ShadowIOleControl, VarPtr(This), PTR_SIZE
    Original_IOleControl_GetControlInfo = ShadowIOleControl.GetControlInfo(CI)
    CopyMemory ShadowIOleControl, NULL_PTR, PTR_SIZE
    This = GetVTableControl()
Else
    Original_IOleControl_GetControlInfo = E_NOTIMPL
End If
End Function

Private Function Original_IOleControl_OnMnemonic(ByRef This As LongPtr, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
If OriginalVTableControl <> NULL_PTR Then
    Dim ShadowIOleControl As OLEGuids.IOleControl
    This = OriginalVTableControl
    CopyMemory ShadowIOleControl, VarPtr(This), PTR_SIZE
    Original_IOleControl_OnMnemonic = ShadowIOleControl.OnMnemonic(Msg)
    CopyMemory ShadowIOleControl, NULL_PTR, PTR_SIZE
    This = GetVTableControl()
Else
    Original_IOleControl_OnMnemonic = E_NOTIMPL
End If
End Function

Private Function Original_IOleControl_OnAmbientPropertyChange(ByRef This As LongPtr, ByVal DispId As Long) As Long
If OriginalVTableControl <> NULL_PTR Then
    Dim ShadowIOleControl As OLEGuids.IOleControl
    This = OriginalVTableControl
    CopyMemory ShadowIOleControl, VarPtr(This), PTR_SIZE
    ShadowIOleControl.OnAmbientPropertyChange DispId
    CopyMemory ShadowIOleControl, NULL_PTR, PTR_SIZE
    This = GetVTableControl()
End If
' This function returns S_OK in all cases.
Original_IOleControl_OnAmbientPropertyChange = S_OK
End Function

Private Function Original_IOleControl_FreezeEvents(ByRef This As LongPtr, ByVal bFreeze As Long) As Long
If OriginalVTableControl <> NULL_PTR Then
    Dim ShadowIOleControl As OLEGuids.IOleControl
    This = OriginalVTableControl
    CopyMemory ShadowIOleControl, VarPtr(This), PTR_SIZE
    ShadowIOleControl.FreezeEvents bFreeze
    CopyMemory ShadowIOleControl, NULL_PTR, PTR_SIZE
    This = GetVTableControl()
End If
' This function returns S_OK in all cases.
Original_IOleControl_FreezeEvents = S_OK
End Function

Private Sub ReplaceIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
If OriginalVTablePPB = NULL_PTR Then CopyMemory OriginalVTablePPB, ByVal ObjPtr(This), PTR_SIZE
If OriginalVTablePPB <> NULL_PTR Then CopyMemory ByVal ObjPtr(This), ByVal VarPtr(GetVTablePPB()), PTR_SIZE
End Sub

Private Sub RestoreIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
If OriginalVTablePPB <> NULL_PTR Then CopyMemory ByVal ObjPtr(This), OriginalVTablePPB, PTR_SIZE
End Sub

Private Function GetVTablePPB() As LongPtr
If VTablePPB(0) = NULL_PTR Then
    If OriginalVTablePPB <> NULL_PTR Then CopyMemory VTablePPB(0), ByVal OriginalVTablePPB, 3 * PTR_SIZE
    VTablePPB(3) = ProcPtr(AddressOf IPPB_GetDisplayString)
    If OriginalVTablePPB <> NULL_PTR Then
        CopyMemory VTablePPB(4), ByVal UnsignedAdd(OriginalVTablePPB, 4 * PTR_SIZE), PTR_SIZE
    Else
        VTablePPB(4) = ProcPtr(AddressOf IPPB_MapPropertyToPage)
    End If
    VTablePPB(5) = ProcPtr(AddressOf IPPB_GetPredefinedStrings)
    VTablePPB(6) = ProcPtr(AddressOf IPPB_GetPredefinedValue)
End If
GetVTablePPB = VarPtr(VTablePPB(0))
End Function

Private Function IPPB_GetDisplayString(ByRef This As LongPtr, ByVal DispId As Long, ByRef lpDisplayName As LongPtr) As Long
If VarPtr(lpDisplayName) = NULL_PTR Then
    IPPB_GetDisplayString = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB, Handled As Boolean, DisplayName As String
Set ShadowIPerPropertyBrowsingVB = PtrToObj(VarPtr(This))
ShadowIPerPropertyBrowsingVB.GetDisplayString Handled, DispId, DisplayName
If Handled = False Then
    IPPB_GetDisplayString = Original_IPPB_GetDisplayString(This, DispId, lpDisplayName)
Else
    lpDisplayName = SysAllocString(StrPtr(DisplayName))
    IPPB_GetDisplayString = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetDisplayString = Original_IPPB_GetDisplayString(This, DispId, lpDisplayName)
End Function

Private Function IPPB_MapPropertyToPage(ByRef This As LongPtr, ByVal DispId As Long, ByRef pCLSID As OLEGuids.OLECLSID) As Long
IPPB_MapPropertyToPage = Original_IPPB_MapPropertyToPage(This, DispId, pCLSID)
End Function

Private Function IPPB_GetPredefinedStrings(ByRef This As LongPtr, ByVal DispId As Long, ByRef pCaStringsOut As OLEGuids.OLECALPOLESTR, ByRef pCaCookiesOut As OLEGuids.OLECADWORD) As Long
If VarPtr(pCaStringsOut) = NULL_PTR Or VarPtr(pCaCookiesOut) = NULL_PTR Then
    IPPB_GetPredefinedStrings = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB, Handled As Boolean
ReDim StringsOutArray(0) As String
ReDim CookiesOutArray(0) As Long
Set ShadowIPerPropertyBrowsingVB = PtrToObj(VarPtr(This))
ShadowIPerPropertyBrowsingVB.GetPredefinedStrings Handled, DispId, StringsOutArray(), CookiesOutArray()
If Handled = False Or UBound(StringsOutArray()) = 0 Then
    IPPB_GetPredefinedStrings = Original_IPPB_GetPredefinedStrings(This, DispId, pCaStringsOut, pCaCookiesOut)
Else
    Dim cElems As Long, pElems As LongPtr, nElemCount As Long
    Dim Buffer As String, lpString As LongPtr
    cElems = UBound(StringsOutArray())
    If Not UBound(CookiesOutArray()) = cElems Then ReDim Preserve CookiesOutArray(cElems) As Long
    pElems = CoTaskMemAlloc(cElems * PTR_SIZE)
    pCaStringsOut.cElems = cElems
    pCaStringsOut.pElems = pElems
    For nElemCount = 0 To cElems - 1
        Buffer = StringsOutArray(nElemCount) & vbNullChar
        lpString = CoTaskMemAlloc(LenB(Buffer))
        CopyMemory ByVal lpString, ByVal StrPtr(Buffer), LenB(Buffer)
        CopyMemory ByVal UnsignedAdd(pElems, nElemCount * PTR_SIZE), ByVal VarPtr(lpString), PTR_SIZE
    Next nElemCount
    pElems = CoTaskMemAlloc(cElems * 4)
    pCaCookiesOut.cElems = cElems
    pCaCookiesOut.pElems = pElems
    For nElemCount = 0 To cElems - 1
        CopyMemory ByVal UnsignedAdd(pElems, nElemCount * 4), CookiesOutArray(nElemCount), 4
    Next nElemCount
    IPPB_GetPredefinedStrings = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetPredefinedStrings = Original_IPPB_GetPredefinedStrings(This, DispId, pCaStringsOut, pCaCookiesOut)
End Function

Private Function IPPB_GetPredefinedValue(ByRef This As LongPtr, ByVal DispId As Long, ByVal dwCookie As Long, ByRef pVarOut As Variant) As Long
If VarPtr(pVarOut) = NULL_PTR Then
    IPPB_GetPredefinedValue = E_POINTER
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB, Handled As Boolean
Set ShadowIPerPropertyBrowsingVB = PtrToObj(VarPtr(This))
ShadowIPerPropertyBrowsingVB.GetPredefinedValue Handled, DispId, dwCookie, pVarOut
If Handled = False Then
    IPPB_GetPredefinedValue = Original_IPPB_GetPredefinedValue(This, DispId, dwCookie, pVarOut)
Else
    IPPB_GetPredefinedValue = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetPredefinedValue = Original_IPPB_GetPredefinedValue(This, DispId, dwCookie, pVarOut)
End Function

Private Function Original_IPPB_GetDisplayString(ByRef This As LongPtr, ByVal DispId As Long, ByRef lpDisplayName As LongPtr) As Long
If OriginalVTablePPB <> NULL_PTR Then
    Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
    This = OriginalVTablePPB
    CopyMemory ShadowIPPB, VarPtr(This), PTR_SIZE
    Original_IPPB_GetDisplayString = ShadowIPPB.GetDisplayString(DispId, lpDisplayName)
    CopyMemory ShadowIPPB, NULL_PTR, PTR_SIZE
    This = GetVTablePPB()
Else
    Original_IPPB_GetDisplayString = E_NOTIMPL
End If
End Function

Private Function Original_IPPB_MapPropertyToPage(ByRef This As LongPtr, ByVal DispId As Long, ByRef pCLSID As OLEGuids.OLECLSID) As Long
If OriginalVTablePPB <> NULL_PTR Then
    Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
    This = OriginalVTablePPB
    CopyMemory ShadowIPPB, VarPtr(This), PTR_SIZE
    Original_IPPB_MapPropertyToPage = ShadowIPPB.MapPropertyToPage(DispId, pCLSID)
    CopyMemory ShadowIPPB, NULL_PTR, PTR_SIZE
    This = GetVTablePPB()
Else
    Original_IPPB_MapPropertyToPage = E_NOTIMPL
End If
End Function

Private Function Original_IPPB_GetPredefinedStrings(ByRef This As LongPtr, ByVal DispId As Long, ByRef pCaStringsOut As OLEGuids.OLECALPOLESTR, ByRef pCaCookiesOut As OLEGuids.OLECADWORD) As Long
If OriginalVTablePPB <> NULL_PTR Then
    Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
    This = OriginalVTablePPB
    CopyMemory ShadowIPPB, VarPtr(This), PTR_SIZE
    Original_IPPB_GetPredefinedStrings = ShadowIPPB.GetPredefinedStrings(DispId, pCaStringsOut, pCaCookiesOut)
    CopyMemory ShadowIPPB, NULL_PTR, PTR_SIZE
    This = GetVTablePPB()
Else
    Original_IPPB_GetPredefinedStrings = E_NOTIMPL
End If
End Function

Private Function Original_IPPB_GetPredefinedValue(ByRef This As LongPtr, ByVal DispId As Long, ByVal dwCookie As Long, ByRef pVarOut As Variant) As Long
If OriginalVTablePPB <> NULL_PTR Then
    Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
    This = OriginalVTablePPB
    CopyMemory ShadowIPPB, VarPtr(This), PTR_SIZE
    Original_IPPB_GetPredefinedValue = ShadowIPPB.GetPredefinedValue(DispId, dwCookie, pVarOut)
    CopyMemory ShadowIPPB, NULL_PTR, PTR_SIZE
    This = GetVTablePPB()
Else
    Original_IPPB_GetPredefinedValue = E_NOTIMPL
End If
End Function

Private Sub ReplaceInPlaceObjectWindowless(ByVal This As OLEGuids.IOleInPlaceObjectWindowless)
If OriginalVTableInPlaceObjectWindowless = NULL_PTR Then CopyMemory OriginalVTableInPlaceObjectWindowless, ByVal ObjPtr(This), PTR_SIZE
If OriginalVTableInPlaceObjectWindowless <> NULL_PTR Then CopyMemory ByVal ObjPtr(This), ByVal VarPtr(GetVTableInPlaceObjectWindowless()), PTR_SIZE
End Sub

Private Sub RestoreInPlaceObjectWindowless(ByVal This As OLEGuids.IOleInPlaceObjectWindowless)
If OriginalVTableInPlaceObjectWindowless <> NULL_PTR Then CopyMemory ByVal ObjPtr(This), OriginalVTableInPlaceObjectWindowless, PTR_SIZE
End Sub

Private Function GetVTableInPlaceObjectWindowless() As LongPtr
If VTableInPlaceObjectWindowless(0) = NULL_PTR Then
    If OriginalVTableInPlaceObjectWindowless <> NULL_PTR Then CopyMemory VTableInPlaceObjectWindowless(0), ByVal OriginalVTableInPlaceObjectWindowless, 9 * PTR_SIZE
    VTableInPlaceObjectWindowless(9) = ProcPtr(AddressOf InPlaceObjectWindowless_OnWindowMessage)
    If OriginalVTableInPlaceObjectWindowless <> NULL_PTR Then
        CopyMemory VTableInPlaceObjectWindowless(10), ByVal UnsignedAdd(OriginalVTableInPlaceObjectWindowless, 10 * PTR_SIZE), PTR_SIZE
    Else
        VTableInPlaceObjectWindowless(10) = ProcPtr(AddressOf InPlaceObjectWindowless_GetDropTarget)
    End If
End If
GetVTableInPlaceObjectWindowless = VarPtr(VTableInPlaceObjectWindowless(0))
End Function

Private Function InPlaceObjectWindowless_OnWindowMessage(ByRef This As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByRef Result As LongPtr) As Long
If OriginalVTableInPlaceObjectWindowless = NULL_PTR Then
    InPlaceObjectWindowless_OnWindowMessage = S_FALSE
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim ShadowIOleInPlaceObjectWindowlessVB As OLEGuids.IOleInPlaceObjectWindowlessVB, Handled As Boolean
Set ShadowIOleInPlaceObjectWindowlessVB = PtrToObj(VarPtr(This))
ShadowIOleInPlaceObjectWindowlessVB.OnWindowMessage Handled, wMsg, wParam, lParam, Result
If Handled = False Then
    InPlaceObjectWindowless_OnWindowMessage = Original_InPlaceObjectWindowless_OnWindowMessage(This, wMsg, wParam, lParam, Result)
Else
    InPlaceObjectWindowless_OnWindowMessage = S_OK
End If
Exit Function
CATCH_EXCEPTION:
InPlaceObjectWindowless_OnWindowMessage = Original_InPlaceObjectWindowless_OnWindowMessage(This, wMsg, wParam, lParam, Result)
End Function

Private Function InPlaceObjectWindowless_GetDropTarget(ByRef This As LongPtr, ByRef lppDropTarget As LongPtr) As Long
InPlaceObjectWindowless_GetDropTarget = Original_InPlaceObjectWindowless_GetDropTarget(This, lppDropTarget)
End Function

Private Function Original_InPlaceObjectWindowless_OnWindowMessage(ByRef This As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByRef Result As LongPtr) As Long
If OriginalVTableInPlaceObjectWindowless <> NULL_PTR Then
    Dim ShadowIOleInPlaceObjectWindowless As OLEGuids.IOleInPlaceObjectWindowless
    This = OriginalVTableInPlaceObjectWindowless
    CopyMemory ShadowIOleInPlaceObjectWindowless, VarPtr(This), PTR_SIZE
    Original_InPlaceObjectWindowless_OnWindowMessage = ShadowIOleInPlaceObjectWindowless.OnWindowMessage(wMsg, wParam, lParam, Result)
    CopyMemory ShadowIOleInPlaceObjectWindowless, NULL_PTR, PTR_SIZE
    This = GetVTableInPlaceObjectWindowless()
Else
    Original_InPlaceObjectWindowless_OnWindowMessage = S_FALSE
End If
End Function

Private Function Original_InPlaceObjectWindowless_GetDropTarget(ByRef This As LongPtr, ByRef lppDropTarget As LongPtr) As Long
If OriginalVTableInPlaceObjectWindowless <> NULL_PTR Then
    Dim ShadowIOleInPlaceObjectWindowless As OLEGuids.IOleInPlaceObjectWindowless
    This = OriginalVTableInPlaceObjectWindowless
    CopyMemory ShadowIOleInPlaceObjectWindowless, VarPtr(This), PTR_SIZE
    Original_InPlaceObjectWindowless_GetDropTarget = ShadowIOleInPlaceObjectWindowless.GetDropTarget(lppDropTarget)
    CopyMemory ShadowIOleInPlaceObjectWindowless, NULL_PTR, PTR_SIZE
    This = GetVTableInPlaceObjectWindowless()
Else
    Original_InPlaceObjectWindowless_GetDropTarget = E_NOTIMPL
End If
End Function

Public Function GetNewEnum(ByVal This As Object, ByVal Upper As Long, ByVal Lower As Long) As IEnumVARIANT
Dim VTableEnumVARIANTData As VTableEnumVARIANTDataStruct
With VTableEnumVARIANTData
.VTable = GetVTableEnumVARIANT()
.RefCount = 1
Set .Enumerable = This
.Index = Lower
.Count = Upper
Dim hMem As LongPtr
hMem = CoTaskMemAlloc(LenB(VTableEnumVARIANTData))
If hMem <> NULL_PTR Then
    CopyMemory ByVal hMem, VTableEnumVARIANTData, LenB(VTableEnumVARIANTData)
    CopyMemory ByVal VarPtr(GetNewEnum), hMem, PTR_SIZE
    CopyMemory ByVal VarPtr(.Enumerable), NULL_PTR, PTR_SIZE
End If
End With
End Function

Private Function GetVTableEnumVARIANT() As LongPtr
If VTableEnumVARIANT(0) = NULL_PTR Then
    VTableEnumVARIANT(0) = ProcPtr(AddressOf IEnumVARIANT_QueryInterface)
    VTableEnumVARIANT(1) = ProcPtr(AddressOf IEnumVARIANT_AddRef)
    VTableEnumVARIANT(2) = ProcPtr(AddressOf IEnumVARIANT_Release)
    VTableEnumVARIANT(3) = ProcPtr(AddressOf IEnumVARIANT_Next)
    VTableEnumVARIANT(4) = ProcPtr(AddressOf IEnumVARIANT_Skip)
    VTableEnumVARIANT(5) = ProcPtr(AddressOf IEnumVARIANT_Reset)
    VTableEnumVARIANT(6) = ProcPtr(AddressOf IEnumVARIANT_Clone)
End If
GetVTableEnumVARIANT = VarPtr(VTableEnumVARIANT(0))
End Function

Private Function IEnumVARIANT_QueryInterface(ByRef This As VTableEnumVARIANTDataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As LongPtr) As Long
If VarPtr(pvObj) = NULL_PTR Then
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

Private Function IEnumVARIANT_AddRef(ByRef This As VTableEnumVARIANTDataStruct) As Long
This.RefCount = This.RefCount + 1
IEnumVARIANT_AddRef = This.RefCount
End Function

Private Function IEnumVARIANT_Release(ByRef This As VTableEnumVARIANTDataStruct) As Long
This.RefCount = This.RefCount - 1
IEnumVARIANT_Release = This.RefCount
If IEnumVARIANT_Release = 0 Then
    Set This.Enumerable = Nothing
    CoTaskMemFree VarPtr(This)
End If
End Function

Private Function IEnumVARIANT_Next(ByRef This As VTableEnumVARIANTDataStruct, ByVal VntCount As Long, ByVal VntArrPtr As LongPtr, ByRef pcvFetched As Long) As Long
If VntArrPtr = NULL_PTR Then
    IEnumVARIANT_Next = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
#If Win64 Then
Const VARIANT_CB As Long = 24
#Else
Const VARIANT_CB As Long = 16
#End If
Dim Fetched As Long
With This
Do Until .Index > .Count
    VariantCopy ByVal VntArrPtr, .Enumerable(.Index)
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
If VarPtr(pcvFetched) <> NULL_PTR Then pcvFetched = Fetched
Exit Function
CATCH_EXCEPTION:
If VarPtr(pcvFetched) <> NULL_PTR Then pcvFetched = 0
IEnumVARIANT_Next = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Skip(ByRef This As VTableEnumVARIANTDataStruct, ByVal VntCount As Long) As Long
IEnumVARIANT_Skip = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Reset(ByRef This As VTableEnumVARIANTDataStruct) As Long
IEnumVARIANT_Reset = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Clone(ByRef This As VTableEnumVARIANTDataStruct, ByRef ppEnum As IEnumVARIANT) As Long
IEnumVARIANT_Clone = E_NOTIMPL
End Function
