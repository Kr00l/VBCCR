Attribute VB_Name = "CommonDialogBase"
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
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type VTablePDEXCallbackDataStruct
VTable As LongPtr
RefCount As Long
ObjectPointer As LongPtr
End Type
#If VBA7 Then
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadID As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As LongPtr
Private Declare PtrSafe Function IsDialogMessage Lib "user32" Alias "IsDialogMessageW" (ByVal hDlg As LongPtr, ByRef lpMsg As TMSG) As Long
Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageW" (ByVal hDlg As Long, ByRef lpMsg As TMSG) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private Const WM_INITDIALOG As Long = &H110
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0
Private CdlSubclassProcPtr As LongPtr
Private CdlPDEXVTableIPDCB(0 To 5) As LongPtr
Private CdlFRHookHandle As LongPtr
Private CdlFRDialogHandle() As LongPtr, CdlFRDialogCount As Long

#If VBA7 Then
Public Sub CdlSetSubclass(ByVal hWnd As LongPtr, ByVal This As CommonDialog, ByVal dwRefData As LongPtr, Optional ByVal Name As String)
#Else
Public Sub CdlSetSubclass(ByVal hWnd As Long, ByVal This As CommonDialog, ByVal dwRefData As Long, Optional ByVal Name As String)
#End If
If hWnd = NULL_PTR Then Exit Sub
If Name = vbNullString Then Name = "Cdl"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If CdlSubclassProcPtr = NULL_PTR Then CdlSubclassProcPtr = ProcPtr(AddressOf CdlSubclassProc)
    SetWindowSubclass hWnd, CdlSubclassProcPtr, ObjPtr(This), dwRefData
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

#If VBA7 Then
Public Function CdlDefaultProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlDefaultProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
CdlDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
End Function

#If VBA7 Then
Public Sub CdlRemoveSubclass(ByVal hWnd As LongPtr, Optional ByVal Name As String)
#Else
Public Sub CdlRemoveSubclass(ByVal hWnd As Long, Optional ByVal Name As String)
#End If
If hWnd = NULL_PTR Then Exit Sub
If Name = vbNullString Then Name = "Cdl"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    RemoveWindowSubclass hWnd, CdlSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

#If VBA7 Then
Public Function CdlSubclassProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Public Function CdlSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
#End If
Select Case wMsg
    Case WM_DESTROY
        CdlSubclassProc = CdlDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        CdlSubclassProc = CdlDefaultProc(hWnd, wMsg, wParam, lParam)
        RemoveWindowSubclass hWnd, CdlSubclassProcPtr, uIdSubclass
        Exit Function
End Select
On Error Resume Next
Dim This As CommonDialog
Set This = PtrToObj(uIdSubclass)
If Err.Number = 0 Then
    CdlSubclassProc = This.FMessage(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    CdlSubclassProc = CdlDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

#If VBA7 Then
Public Function CdlOFN1CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlOFN1CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("CdlOFN1CallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("CdlOFN1CallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As CommonDialog
    Set This = PtrToObj(lCustData)
    CdlOFN1CallbackProc = This.FMessage(hDlg, wMsg, wParam, lParam, -1)
Else
    CdlOFN1CallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function CdlOFN1CallbackProcOldStyle(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlOFN1CallbackProcOldStyle(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("CdlOFN1CallbackProcOldStyleCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("CdlOFN1CallbackProcOldStyleCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As CommonDialog
    Set This = PtrToObj(lCustData)
    CdlOFN1CallbackProcOldStyle = This.FMessage(hDlg, wMsg, wParam, lParam, -1001)
Else
    CdlOFN1CallbackProcOldStyle = 0
End If
End Function

#If VBA7 Then
Public Function CdlOFN2CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlOFN2CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("CdlOFN2CallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("CdlOFN2CallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As CommonDialog
    Set This = PtrToObj(lCustData)
    CdlOFN2CallbackProc = This.FMessage(hDlg, wMsg, wParam, lParam, -2)
Else
    CdlOFN2CallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function CdlOFN2CallbackProcOldStyle(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlOFN2CallbackProcOldStyle(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("CdlOFN2CallbackProcOldStyleCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 112), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("CdlOFN2CallbackProcOldStyleCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As CommonDialog
    Set This = PtrToObj(lCustData)
    CdlOFN2CallbackProcOldStyle = This.FMessage(hDlg, wMsg, wParam, lParam, -1002)
Else
    CdlOFN2CallbackProcOldStyle = 0
End If
End Function

#If VBA7 Then
Public Function CdlCCCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlCCCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("CdlCCCallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 48), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 24), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("CdlCCCallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As CommonDialog
    Set This = PtrToObj(lCustData)
    CdlCCCallbackProc = This.FMessage(hDlg, wMsg, wParam, lParam, -3)
Else
    CdlCCCallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function CdlCFCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlCFCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
If wMsg <> WM_INITDIALOG Then
    lCustData = GetProp(hDlg, StrPtr("CdlCFCallbackProcCustData"))
Else
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 40), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), PTR_SIZE
    #End If
    SetProp hDlg, StrPtr("CdlCFCallbackProcCustData"), lCustData
End If
If lCustData <> NULL_PTR Then
    Dim This As CommonDialog
    Set This = PtrToObj(lCustData)
    CdlCFCallbackProc = This.FMessage(hDlg, wMsg, wParam, lParam, -4)
Else
    CdlCFCallbackProc = 0
End If
End Function

#If VBA7 Then
Public Function CdlPDCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlPDCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    CdlPDCallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 54), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 38), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As CommonDialog
        Set This = PtrToObj(lCustData)
        CdlPDCallbackProc = This.FMessage(hDlg, wMsg, wParam, lParam, -5)
    Else
        CdlPDCallbackProc = 0
    End If
End If
End Function

#If VBA7 Then
Public Function CdlPDEXCallbackPtr(ByVal This As CommonDialog) As LongPtr
#Else
Public Function CdlPDEXCallbackPtr(ByVal This As CommonDialog) As Long
#End If
Dim VTableData(0 To 2) As LongPtr
VTableData(0) = GetVTableIPDCB()
VTableData(1) = 0 ' RefCount is uninstantiated
VTableData(2) = ObjPtr(This)
Dim hMem As LongPtr
hMem = CoTaskMemAlloc(3 * PTR_SIZE)
If hMem <> NULL_PTR Then
    CopyMemory ByVal hMem, VTableData(0), 3 * PTR_SIZE
    CdlPDEXCallbackPtr = hMem
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
        Dim This As CommonDialog
        Set This = PtrToObj(lCustData)
        This.FMessage hDlg, wMsg, wParam, lParam, -5
    End If
End If
IPDCB_HandleMessage = S_FALSE
End Function

#If VBA7 Then
Public Function CdlPSDCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlPSDCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    CdlPSDCallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 88), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 64), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As CommonDialog
        Set This = PtrToObj(lCustData)
        CdlPSDCallbackProc = This.FMessage(hDlg, wMsg, wParam, lParam, -7)
    Else
        CdlPSDCallbackProc = 0
    End If
End If
End Function

#If VBA7 Then
Public Function CdlBIFCallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal lParam As LongPtr, ByVal This As CommonDialog) As Long
#Else
Public Function CdlBIFCallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal lParam As Long, ByVal This As CommonDialog) As Long
#End If
CdlBIFCallbackProc = CLng(This.FMessage(hDlg, wMsg, 0, lParam, -8))
End Function

#If VBA7 Then
Public Function CdlFR1CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlFR1CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    CdlFR1CallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 56), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As CommonDialog
        Set This = PtrToObj(lCustData)
        This.FMessage hDlg, wMsg, wParam, lParam, -9
    End If
    ' Need to return a nonzero value or else the dialog box will not be shown.
    CdlFR1CallbackProc = 1
End If
End Function

#If VBA7 Then
Public Function CdlFR2CallbackProc(ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function CdlFR2CallbackProc(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
If wMsg <> WM_INITDIALOG Then
    CdlFR2CallbackProc = 0
Else
    Dim lCustData As LongPtr
    #If Win64 Then
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 56), PTR_SIZE
    #Else
    CopyMemory lCustData, ByVal UnsignedAdd(lParam, 28), PTR_SIZE
    #End If
    If lCustData <> NULL_PTR Then
        Dim This As CommonDialog
        Set This = PtrToObj(lCustData)
        This.FMessage hDlg, wMsg, wParam, lParam, -10
    End If
    ' Need to return a nonzero value or else the dialog box will not be shown.
    CdlFR2CallbackProc = 1
End If
End Function

#If VBA7 Then
Public Sub CdlFRAddHook(ByVal hDlg As LongPtr)
#Else
Public Sub CdlFRAddHook(ByVal hDlg As Long)
#End If
If CdlFRHookHandle = NULL_PTR And CdlFRDialogCount = 0 Then
    Const WH_GETMESSAGE As Long = 3
    CdlFRHookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf CdlFRHookProc, 0, App.ThreadID)
    ReDim CdlFRDialogHandle(0) ' As LongPtr
    CdlFRDialogHandle(0) = hDlg
Else
    ReDim Preserve CdlFRDialogHandle(0 To CdlFRDialogCount) ' As LongPtr
    CdlFRDialogHandle(CdlFRDialogCount) = hDlg
End If
CdlFRDialogCount = CdlFRDialogCount + 1
End Sub

#If VBA7 Then
Public Sub CdlFRReleaseHook(ByVal hDlg As LongPtr)
#Else
Public Sub CdlFRReleaseHook(ByVal hDlg As Long)
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

Private Function CdlFRHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
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
CdlFRHookProc = CallNextHookEx(CdlFRHookHandle, nCode, wParam, lParam)
End Function

Private Function UnsignedAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr
#If Win64 Then
UnsignedAdd = ((Start Xor &H8000000000000000^) + Incr) Xor &H8000000000000000^
#Else
UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
#End If
End Function

Private Function ProcPtr(ByVal Address As LongPtr) As LongPtr
ProcPtr = Address
End Function

Private Function PtrToObj(ByVal ObjectPointer As LongPtr) As Object
Dim TempObj As Object
CopyMemory TempObj, ObjectPointer, PTR_SIZE
Set PtrToObj = TempObj
CopyMemory TempObj, NULL_PTR, PTR_SIZE
End Function
