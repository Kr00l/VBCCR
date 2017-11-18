Attribute VB_Name = "RichTextBoxBase"
Option Explicit

' Required:

' OLEGuids.tlb (in IDE only)

Private Enum VTableIndexRichEditOleCallbackConstants
' Ignore : RichEditOleCallbackQueryInterface
' Ignore : RichEditOleCallbackAddRef
' Ignore : RichEditOleCallbackRelease
VTableIndexRichEditOleCallbackGetNewStorage = 4
VTableIndexRichEditOleCallbackGetInPlaceContext = 5
VTableIndexRichEditOleCallbackShowContainerUI = 6
VTableIndexRichEditOleCallbackQueryInsertObject = 7
VTableIndexRichEditOleCallbackDeleteObject = 8
VTableIndexRichEditOleCallbackQueryAcceptData = 9
VTableIndexRichEditOleCallbackContextSensitiveHelp = 10
VTableIndexRichEditOleCallbackGetClipboardData = 11
VTableIndexRichEditOleCallbackGetDragDropEffect = 12
VTableIndexRichEditOleCallbackGetContextMenu = 13
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Const E_NOTIMPL As Long = &H80004001
Private Const S_OK As Long = &H0
Private RichedModHandle As Long, RichedModCount As Long, RichedClassName As String
Private StreamStringOut() As Byte, StreamStringOutUBound As Long
Private StreamStringIn() As Byte, StreamStringInLength As Long, StreamStringInPos As Long
Private VTableSubclassRichEditOleCallback As VTableSubclass

Public Sub RtfLoadRichedMod()
If (RichedModHandle Or RichedModCount) = 0 Then
    RichedModHandle = LoadLibrary(StrPtr("Msftedit.dll"))
    If RichedModHandle <> 0 Then
        RichedClassName = "RichEdit50W"
    Else
        RichedModHandle = LoadLibrary(StrPtr("Riched20.dll"))
        RichedClassName = "RichEdit20W"
    End If
End If
RichedModCount = RichedModCount + 1
End Sub

Public Sub RtfReleaseRichedMod()
RichedModCount = RichedModCount - 1
If RichedModCount = 0 And RichedModHandle <> 0 Then
    FreeLibrary RichedModHandle
    RichedModHandle = 0
End If
End Sub

Public Function RtfGetClassName() As String
RtfGetClassName = RichedClassName
End Function

Public Function RtfStreamStringOut() As String
RtfStreamStringOut = StreamStringOut()
Erase StreamStringOut()
StreamStringOutUBound = 0
End Function

Public Function RtfStreamCallbackStringOut(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
If BytesRequested > 0 Then
    ReDim Preserve StreamStringOut(0 To (StreamStringOutUBound + BytesRequested - 1)) As Byte
    CopyMemory StreamStringOut(StreamStringOutUBound), ByVal ByteBufferPtr, BytesRequested
    StreamStringOutUBound = StreamStringOutUBound + BytesRequested
    BytesProcessed = BytesRequested
Else
    BytesProcessed = 0
End If
RtfStreamCallbackStringOut = 0
End Function

Public Sub RtfStreamStringIn(ByVal Value As String)
StreamStringInLength = LenB(Value)
Erase StreamStringIn()
If StreamStringInLength > 0 Then
    ReDim StreamStringIn(0 To (StreamStringInLength - 1)) As Byte
    CopyMemory StreamStringIn(0), ByVal StrPtr(Value), StreamStringInLength
End If
StreamStringInPos = 0
End Sub

Public Sub RtfStreamStringInCleanUp()
Erase StreamStringIn()
StreamStringInLength = 0
StreamStringInPos = 0
End Sub

Public Function RtfStreamCallbackStringIn(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
If BytesRequested > (StreamStringInLength - StreamStringInPos) Then BytesRequested = (StreamStringInLength - StreamStringInPos)
If BytesRequested > 0 Then
    CopyMemory ByVal ByteBufferPtr, StreamStringIn(StreamStringInPos), BytesRequested
    StreamStringInPos = StreamStringInPos + BytesRequested
Else
    BytesRequested = 0
End If
BytesProcessed = BytesRequested
RtfStreamCallbackStringIn = 0
End Function

Public Function RtfStreamCallbackFileOut(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
RtfStreamCallbackFileOut = IIf(WriteFile(dwCookie, ByteBufferPtr, BytesRequested, BytesProcessed, 0) <> 0, 0, 1)
End Function

Public Function RtfStreamCallbackFileIn(ByVal dwCookie As Long, ByVal ByteBufferPtr As Long, ByVal BytesRequested As Long, ByRef BytesProcessed As Long) As Long
RtfStreamCallbackFileIn = IIf(ReadFile(dwCookie, ByteBufferPtr, BytesRequested, BytesProcessed, 0) <> 0, 0, 1)
End Function

Public Sub SetVTableSubclassIRichEditOleCallback(ByVal This As Object)
If VTableSupported(This) = True Then
    Dim ShadowIRichEditOleCallback As OLEGuids.IRichEditOleCallback
    Set ShadowIRichEditOleCallback = This
    Call ReplaceIRichEditOleCallback(This)
End If
End Sub

Public Sub RemoveVTableSubclassIRichEditOleCallback(ByVal This As Object)
Attribute RemoveVTableSubclassIRichEditOleCallback.VB_MemberFlags = "40"
If VTableSupported(This) = True Then Call RestoreIRichEditOleCallback(This)
End Sub

Private Function VTableSupported(ByRef This As Object) As Boolean
On Error GoTo Cancel
Dim ShadowIRichEditOleCallback As OLEGuids.IRichEditOleCallback
Set ShadowIRichEditOleCallback = This
VTableSupported = Not CBool(ShadowIRichEditOleCallback Is Nothing)
Cancel:
End Function

Private Sub ReplaceIRichEditOleCallback(ByVal This As OLEGuids.IRichEditOleCallback)
If VTableSubclassRichEditOleCallback Is Nothing Then Set VTableSubclassRichEditOleCallback = New VTableSubclass
If VTableSubclassRichEditOleCallback.RefCount = 0 Then
    VTableSubclassRichEditOleCallback.Subclass ObjPtr(This), VTableIndexRichEditOleCallbackGetNewStorage, VTableIndexRichEditOleCallbackGetContextMenu, _
    AddressOf IRichEditOleCallback_GetNewStorage, AddressOf IRichEditOleCallback_GetInPlaceContext, _
    AddressOf IRichEditOleCallback_ShowContainerUI, AddressOf IRichEditOleCallback_QueryInsertObject, _
    AddressOf IRichEditOleCallback_DeleteObject, AddressOf IRichEditOleCallback_QueryAcceptData, _
    AddressOf IRichEditOleCallback_ContextSensitiveHelp, AddressOf IRichEditOleCallback_GetClipboardData, _
    AddressOf IRichEditOleCallback_GetDragDropEffect, AddressOf IRichEditOleCallback_GetContextMenu
End If
VTableSubclassRichEditOleCallback.AddRef
End Sub

Private Sub RestoreIRichEditOleCallback(ByVal This As OLEGuids.IRichEditOleCallback)
If Not VTableSubclassRichEditOleCallback Is Nothing Then
    VTableSubclassRichEditOleCallback.Release
    If VTableSubclassRichEditOleCallback.RefCount = 0 Then VTableSubclassRichEditOleCallback.UnSubclass
End If
End Sub

Private Function IRichEditOleCallback_GetNewStorage(ByVal This As Object, ByRef ppStorage As OLEGuids.IStorage) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRtfOleCallback As RtfOleCallback
Set ShadowRtfOleCallback = This
ShadowRtfOleCallback.ShadowRichTextBox.FIRichEditOleCallback_GetNewStorage IRichEditOleCallback_GetNewStorage, ppStorage
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_GetNewStorage = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetInPlaceContext(ByVal This As Object, ByRef ppFrame As OLEGuids.IOleInPlaceFrame, ByRef ppDoc As OLEGuids.IOleInPlaceUIWindow, ByRef pFrameInfo As OLEGuids.OLEINPLACEFRAMEINFO) As Long
IRichEditOleCallback_GetInPlaceContext = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_ShowContainerUI(ByVal This As Object, ByVal fShow As Long) As Long
IRichEditOleCallback_ShowContainerUI = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_QueryInsertObject(ByVal This As Object, ByRef pCLSID As OLEGuids.OLECLSID, ByVal pStorage As OLEGuids.IStorage, ByVal CharPos As Long) As Long
IRichEditOleCallback_QueryInsertObject = S_OK
End Function

Private Function IRichEditOleCallback_DeleteObject(ByVal This As Object, ByVal LpOleObject As Long) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRtfOleCallback As RtfOleCallback
Set ShadowRtfOleCallback = This
ShadowRtfOleCallback.ShadowRichTextBox.FIRichEditOleCallback_DeleteObject LpOleObject
IRichEditOleCallback_DeleteObject = S_OK
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_DeleteObject = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_QueryAcceptData(ByVal This As Object, ByVal pDataObject As OLEGuids.IDataObject, ByRef CF As Integer, ByVal RECO As Long, ByVal fReally As Long, ByVal hMetaPict As Long) As Long
IRichEditOleCallback_QueryAcceptData = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_ContextSensitiveHelp(ByVal This As Object, ByVal fEnterMode As Long) As Long
IRichEditOleCallback_ContextSensitiveHelp = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetClipboardData(ByVal This As Object, ByVal lpCharRange As Long, ByVal RECO As Long, ByRef ppDataObject As OLEGuids.IDataObject) As Long
IRichEditOleCallback_GetClipboardData = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetDragDropEffect(ByVal This As Object, ByVal fDrag As Long, ByVal KeyState As Long, ByRef dwEffect As Long) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRtfOleCallback As RtfOleCallback
Set ShadowRtfOleCallback = This
ShadowRtfOleCallback.ShadowRichTextBox.FIRichEditOleCallback_GetDragDropEffect CBool(fDrag <> 0), KeyState, dwEffect
IRichEditOleCallback_GetDragDropEffect = S_OK
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_GetDragDropEffect = E_NOTIMPL
End Function

Private Function IRichEditOleCallback_GetContextMenu(ByVal This As Object, ByVal SelType As Integer, ByVal LpOleObject As Long, ByVal lpCharRange As Long, ByRef hMenu As Long) As Long
On Error GoTo CATCH_EXCEPTION
Dim ShadowRtfOleCallback As RtfOleCallback
Set ShadowRtfOleCallback = This
ShadowRtfOleCallback.ShadowRichTextBox.FIRichEditOleCallback_GetContextMenu SelType, LpOleObject, lpCharRange, hMenu
If hMenu = 0 Then
    IRichEditOleCallback_GetContextMenu = E_NOTIMPL
Else
    IRichEditOleCallback_GetContextMenu = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IRichEditOleCallback_GetContextMenu = E_NOTIMPL
End Function
