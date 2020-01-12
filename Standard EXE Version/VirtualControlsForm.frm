VERSION 5.00
Begin VB.Form VirtualControlsForm 
   Caption         =   "VirtualControls Demo"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin ComCtlsDemo.ListView ListView1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5318
      View            =   3
      AllowColumnReorder=   -1  'True
      MultiSelect     =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      Checkboxes      =   -1  'True
      ShowInfoTips    =   -1  'True
      VirtualMode     =   -1  'True
   End
End
Attribute VB_Name = "VirtualControlsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type VirtualLvwItemStruct
Text As String
Icon As Long
ToolTipText As String
Bold As Boolean
ForeColor As OLE_COLOR
Checked As Boolean
End Type
Private VirtualLvwItems(1 To 100000, 0 To 3) As VirtualLvwItemStruct

Private Sub Form_Load()
Call SetupVisualStyleFixes(Me)
ListView1.VirtualDisabledInfos = 0 ' None disabled info
Set ListView1.SmallIcons = MainForm.ImageList1
Dim i As Long, j As Long
For i = 1 To 100000
    For j = 0 To 3
        With VirtualLvwItems(i, j)
        .ForeColor = -1
        If j = 0 Then
            .Text = "item" & i
            .Icon = 1
            .ToolTipText = "Info " & CStr(i)
        Else
            .Text = "sub text" & j & "_" & IIf(i Mod 2, "B", "A")
            If j = 1 Then
                .ToolTipText = "SubInfo " & CStr(i)
            ElseIf j = 2 Then
                .Icon = 1
                .ForeColor = vbBlue
            Else
                .Bold = True
            End If
        End If
        End With
    Next j
Next i
With ListView1.ColumnHeaders
.Add , , "Col1"
.Add , , "Col2"
.Add , , "Col3"
.Add , , "Col4"
End With
ListView1.VirtualItemCount = 100000
End Sub

Private Sub ListView1_FindVirtualItem(ByVal StartIndex As Long, ByVal SearchText As String, ByVal Partial As Boolean, ByVal Wrap As Boolean, FoundIndex As Long)
' This event must be handled to enable incremental search on key presses
If Count = 0 Then Exit Sub
Dim i As Long
For i = StartIndex To ListView1.VirtualItemCount
    If StrComp(Left$(VirtualLvwItems(i, 0).Text, Len(SearchText)), SearchText, vbTextCompare) = 0 Then
        FoundIndex = i
        Exit For
    End If
Next i
If FoundIndex = 0 And Wrap = True Then
    For i = 1 To StartIndex - 1
        If StrComp(Left$(VirtualLvwItems(i, 0).Text, Len(SearchText)), SearchText, vbTextCompare) = 0 Then
            FoundIndex = i
            Exit For
        End If
    Next i
End If
End Sub

Private Sub ListView1_GetVirtualItem(ByVal Index As Long, ByVal SubItemIndex As Long, ByVal VirtualProperty As LvwVirtualPropertyConstants, Value As Variant)
With VirtualLvwItems(Index, SubItemIndex)
Select Case VirtualProperty
    Case LvwVirtualPropertyText
        Value = .Text
    Case LvwVirtualPropertyIcon
        Value = .Icon
    Case LvwVirtualPropertyIndentation
        Value = 0
    Case LvwVirtualPropertyToolTipText
        Value = .ToolTipText
    Case LvwVirtualPropertyBold
        Value = .Bold
    Case LvwVirtualPropertyForeColor
        If .ForeColor <> -1 Then Value = .ForeColor
    Case LvwVirtualPropertyChecked
        Value = .Checked
End Select
End With
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As LvwListItem, ByVal Checked As Boolean)
VirtualLvwItems(Item.Index, 0).Checked = Checked
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Boolean, NewString As String)
Dim ListItem As LvwListItem
Set ListItem = ListView1.SelectedItem
If Not ListItem Is Nothing Then VirtualLvwItems(ListItem.Index, 0).Text = NewString
End Sub
