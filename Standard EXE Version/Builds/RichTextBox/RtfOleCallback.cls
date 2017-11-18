VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RtfOleCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements OLEGuids.IRichEditOleCallback
Private ShadowObjectPointer As Long

Private Sub IRichEditOleCallback_GetNewStorage(ByRef ppStorage As OLEGuids.IStorage)
End Sub

Private Sub IRichEditOleCallback_GetInPlaceContext(ByRef ppFrame As OLEGuids.IOleInPlaceFrame, ByRef ppDoc As OLEGuids.IOleInPlaceUIWindow, ByRef pFrameInfo As OLEGuids.OLEINPLACEFRAMEINFO)
End Sub

Private Sub IRichEditOleCallback_ShowContainerUI(ByVal fShow As Long)
End Sub

Private Sub IRichEditOleCallback_QueryInsertObject(ByRef pCLSID As OLEGuids.OLECLSID, ByVal pStorage As OLEGuids.IStorage, ByVal CharPos As Long)
End Sub

Private Sub IRichEditOleCallback_DeleteObject(ByVal LpOleObject As Long)
End Sub

Private Sub IRichEditOleCallback_QueryAcceptData(ByVal pDataObject As OLEGuids.IDataObject, ByRef CF As Integer, ByVal RECO As Long, ByVal fReally As Long, ByVal hMetaPict As Long)
End Sub

Private Sub IRichEditOleCallback_ContextSensitiveHelp(ByVal fEnterMode As Long)
End Sub

Private Sub IRichEditOleCallback_GetClipboardData(ByVal lpCharRange As Long, ByVal RECO As Long, ByRef ppDataObject As OLEGuids.IDataObject)
End Sub

Private Sub IRichEditOleCallback_GetDragDropEffect(ByVal fDrag As Long, ByVal KeyState As Long, ByRef dwEffect As Long)
End Sub

Private Sub IRichEditOleCallback_GetContextMenu(ByVal SelType As Integer, ByVal LpOleObject As Long, ByVal lpCharRange As Long, ByRef hMenu As Long)
End Sub

Public Function ShadowRichTextBox() As RichTextBox
Set ShadowRichTextBox = PtrToObj(ShadowObjectPointer)
End Function

Friend Sub FInit(ByVal Object As RichTextBox)
ShadowObjectPointer = ObjPtr(Object)
End Sub
