VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VTableSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private PropRefCount As Long
Private PropVTableCount As Long
Private VTableHeaderPointer As Long
Private VTable() As Long, VTableOld() As Long

Private Sub Class_Terminate()
If VTableHeaderPointer <> 0 Then Call UnSubclass
End Sub

Public Property Get RefCount() As Long
RefCount = PropRefCount
End Property

Public Sub AddRef()
PropRefCount = PropRefCount + 1
End Sub

Public Sub Release()
PropRefCount = PropRefCount - 1
End Sub

Public Sub Subclass(ByVal ObjectPointer As Long, ByVal FirstEntry As Long, ByVal LastEntry As Long, ParamArray NewEntries() As Variant)
FirstEntry = FirstEntry - 1
Debug.Assert Not (FirstEntry < 0 Or FirstEntry > LastEntry Or LastEntry < 0 Or VTableHeaderPointer <> 0 Or ObjectPointer = 0)
CopyMemory VTableHeaderPointer, ByVal ObjectPointer, 4
PropVTableCount = LastEntry
ReDim VTable(0 To PropVTableCount) As Long
ReDim VTableOld(0 To PropVTableCount) As Long
Dim Entry As Long, EntryPointer As Long
Entry = UBound(NewEntries()) + FirstEntry
If Entry > PropVTableCount Then Entry = PropVTableCount
EntryPointer = UnsignedAdd(VTableHeaderPointer, FirstEntry * 4)
For Entry = FirstEntry To Entry
    VTable(Entry) = NewEntries(Entry - FirstEntry)
    If VTable(Entry) <> 0 Then
        Call CreateSubclass(EntryPointer, VTable(Entry), VTableOld(Entry))
    End If
    EntryPointer = UnsignedAdd(EntryPointer, 4)
Next Entry
End Sub

Public Property Get SubclassEntry(ByVal Entry As Long) As Boolean
Entry = Entry - 1
Debug.Assert Entry > -1 And Entry < PropVTableCount And VTableHeaderPointer <> 0
SubclassEntry = CBool(VTableOld(Entry))
End Property

Public Property Let SubclassEntry(ByVal Entry As Long, ByVal Value As Boolean)
Entry = Entry - 1
Dim EntryPointer As Long
Debug.Assert Entry >= 0 And Entry <= PropVTableCount And VTableHeaderPointer <> 0
If Me.SubclassEntry(Entry + 1) Xor Value Then
    EntryPointer = UnsignedAdd(VTableHeaderPointer, Entry * 4)
    If Value = True Then
        Call CreateSubclass(EntryPointer, VTable(Entry), VTableOld(Entry))
    Else
        Call CreateSubclass(EntryPointer, VTableOld(Entry), 0)
        VTableOld(Entry) = 0
    End If
End If
End Property

Public Sub ReSubclass()
If VTableHeaderPointer <> 0 Then
    Dim i As Long
    For i = 0 To PropVTableCount
        If VTableOld(i) <> 0 Then
            Call CreateSubclass(UnsignedAdd(VTableHeaderPointer, i * 4), VTableOld(i), 0)
            VTableOld(i) = 0
        End If
    Next i
    For i = 0 To PropVTableCount
        If VTable(i) <> 0 Then
            Call CreateSubclass(UnsignedAdd(VTableHeaderPointer, i * 4), VTable(i), VTableOld(i))
        End If
    Next i
End If
End Sub

Public Sub UnSubclass()
If VTableHeaderPointer <> 0 Then
    Dim i As Long
    For i = 0 To PropVTableCount
        If VTableOld(i) <> 0 Then
            Call CreateSubclass(UnsignedAdd(VTableHeaderPointer, i * 4), VTableOld(i), 0)
            VTableOld(i) = 0
        End If
    Next i
    VTableHeaderPointer = 0
End If
End Sub

Private Sub CreateSubclass(ByVal EntryPointer As Long, ByVal NewPointer As Long, ByRef OldPointer As Long)
CopyMemory OldPointer, ByVal EntryPointer, 4
If OldPointer <> NewPointer Then
    Dim OldProtect As Long
    VirtualProtect EntryPointer, 4, PAGE_EXECUTE_READWRITE, OldProtect
    CopyMemory ByVal EntryPointer, NewPointer, 4
    VirtualProtect EntryPointer, 4, OldProtect, OldProtect
Else
    ' If you get this Assert then better restart the IDE.
    ' Known reasons:
    ' - End button was pushed.
    ' - Object has been modified while it is subclassed.
    Debug.Assert CBool(OldPointer <> NewPointer)
End If
End Sub
