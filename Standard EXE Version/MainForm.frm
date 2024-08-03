VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "VBFlexGrid Demo"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13830
   KeyPreview      =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   13830
   Begin VB.PictureBox PicturePanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   13830
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5895
      Width           =   13830
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   9480
         ScaleHeight     =   255
         ScaleWidth      =   1275
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Paste"
         Height          =   315
         Left            =   10200
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Partial Search"
         Height          =   255
         Left            =   10920
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command21 
         Caption         =   "DragRowCol"
         Height          =   315
         Left            =   12360
         TabIndex        =   32
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copy"
         Height          =   315
         Left            =   9480
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command16 
         Caption         =   "FindItem"
         Height          =   315
         Left            =   10920
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         Caption         =   "RowHidden"
         Height          =   315
         Left            =   12360
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Show Property Pages"
         Height          =   315
         Left            =   9480
         TabIndex        =   23
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Printscreen To Clipboard"
         Height          =   315
         Left            =   11640
         TabIndex        =   24
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Open UserEditing Demo (in-cell editing)"
         Height          =   315
         Left            =   9480
         TabIndex        =   25
         Top             =   480
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sorting"
         Height          =   1455
         Left            =   6840
         TabIndex        =   18
         Top             =   120
         Width           =   2415
         Begin VB.CommandButton Command18 
            Caption         =   "Sort Desc"
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Sort Asc"
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "SortType"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cell"
         Height          =   1455
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   4095
         Begin VB.CommandButton Command3 
            Caption         =   "BackColor"
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            Caption         =   "ForeColor"
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Text"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command10 
            Caption         =   "EnsureVisible"
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command11 
            Caption         =   "ClearContent"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton Command12 
            Caption         =   "ClearAll"
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton Command13 
            Caption         =   "ToolTipText"
            Height          =   315
            Left            =   2760
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Font"
            Height          =   315
            Left            =   2760
            TabIndex        =   14
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command15 
            Caption         =   "ClearFont"
            Height          =   315
            Left            =   2760
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CellPicture"
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2415
         Begin VB.CommandButton Command1 
            Caption         =   "Set"
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Clear"
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   960
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   375
            Left            =   120
            Picture         =   "MainForm.frx":0000
            ScaleHeight     =   315
            ScaleWidth      =   435
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "CellPictureAlignment"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   2175
         End
      End
   End
   Begin VBFlexGridDemo.VBFlexGrid VBFlexGrid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9975
      OLEDropMode     =   1
      Rows            =   150
      Cols            =   20
      AllowUserFreezing=   3
      AllowUserResizing=   3
      ShowInfoTips    =   -1  'True
      ShowLabelTips   =   -1  'True
      ShowScrollTips  =   -1  'True
      AllowScrollLock =   -1  'True
      WallPaper       =   "MainForm.frx":0344
      AllowReaderMode =   -1  'True
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
#If VBA7 Then
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Function OleCreatePropertyFrame Lib "oleaut32" (ByVal hWndOwner As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal lpszCaption As LongPtr, ByVal cObjects As Long, ByRef pUnk As IUnknown, ByVal cPages As Long, ByRef pPageCLSID As Any, ByVal LCID As Long, ByVal dwReserved As Long, ByVal pvReserved As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As LongPtr, ByRef pCLSID As Any) As Long
#Else
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function OleCreatePropertyFrame Lib "oleaut32" (ByVal hWndOwner As Long, ByVal X As Long, ByVal Y As Long, ByVal lpszCaption As Long, ByVal cObjects As Long, ByRef pUnk As IUnknown, ByVal cPages As Long, ByRef pPageCLSID As Any, ByVal LCID As Long, ByVal dwReserved As Long, ByVal pvReserved As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
#End If
Private Const CLSID_StandardColorPage As String = "{7EBDAAE1-8120-11CF-899F-00AA00688B10}"
Private Const CLSID_StandardFontPage As String = "{7EBDAAE0-8120-11CF-899F-00AA00688B10}"
Private PropCellBackColor As OLE_COLOR, PropCellForeColor As OLE_COLOR
Private PropCellFont As StdFont
Attribute PropCellFont.VB_VarHelpID = -1
Private PropDragRowColActive As Boolean
Private PropInsertRowDragging As Boolean

Public Property Get CellBackColor() As OLE_COLOR
CellBackColor = PropCellBackColor
End Property

Public Property Let CellBackColor(ByVal Value As OLE_COLOR)
PropCellBackColor = Value
VBFlexGrid1.CellBackColor = PropCellBackColor
End Property

Public Property Get CellForeColor() As OLE_COLOR
CellForeColor = PropCellForeColor
End Property

Public Property Let CellForeColor(ByVal Value As OLE_COLOR)
PropCellForeColor = Value
VBFlexGrid1.CellForeColor = PropCellForeColor
End Property

Public Property Get CellFont() As StdFont
Set CellFont = PropCellFont
End Property

Public Property Let CellFont(ByVal NewFont As StdFont)
Set Me.CellFont = NewFont
End Property

Public Property Set CellFont(ByVal NewFont As StdFont)
Set PropCellFont = NewFont
If PropCellFont Is Nothing Then
    VBFlexGrid1.CellFontName = vbNullString
Else
    VBFlexGrid1.CellFontName = PropCellFont.Name
    VBFlexGrid1.CellFontSize = PropCellFont.Size
    VBFlexGrid1.CellFontBold = PropCellFont.Bold
    VBFlexGrid1.CellFontItalic = PropCellFont.Italic
    VBFlexGrid1.CellFontStrikeThrough = PropCellFont.Strikethrough
    VBFlexGrid1.CellFontUnderline = PropCellFont.Underline
    VBFlexGrid1.CellFontCharset = PropCellFont.Charset
End If
End Property

Private Sub Form_Load()
Call SetupVisualStylesFixes(Me)
Dim i As Long, j As Long, DecStr As String, StartDate As Date
DecStr = Mid$(1.1, 2, 1)
StartDate = DateSerial(Year(Now()), 1, 1)
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1
    For j = VBFlexGrid1.FixedCols To VBFlexGrid1.Cols - 1
        If j <> 1 Then
            VBFlexGrid1.TextMatrix(i, j) = i & DecStr & j
            VBFlexGrid1.Cell(FlexCellToolTipText, i, j) = i & "/" & j & " info tip."
        Else
            VBFlexGrid1.TextMatrix(i, j) = StartDate + (i - 1)
        End If
    Next j
Next i
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1
    VBFlexGrid1.TextMatrix(i, 0) = i
Next i
For j = VBFlexGrid1.FixedCols To VBFlexGrid1.Cols - 1
    VBFlexGrid1.TextMatrix(0, j) = Chr(64 + j)
Next j
With Combo1
.AddItem FlexPictureAlignmentLeftTop & " - LeftTop"
.ItemData(.NewIndex) = FlexPictureAlignmentLeftTop
.AddItem FlexPictureAlignmentLeftCenter & " - LeftCenter"
.ItemData(.NewIndex) = FlexPictureAlignmentLeftCenter
.AddItem FlexPictureAlignmentLeftBottom & " - LeftBottom"
.ItemData(.NewIndex) = FlexPictureAlignmentLeftBottom
.AddItem FlexPictureAlignmentCenterTop & " - CenterTop"
.ItemData(.NewIndex) = FlexPictureAlignmentCenterTop
.AddItem FlexPictureAlignmentCenterCenter & " - CenterCenter"
.ItemData(.NewIndex) = FlexPictureAlignmentCenterCenter
.AddItem FlexPictureAlignmentCenterBottom & " - CenterBottom"
.ItemData(.NewIndex) = FlexPictureAlignmentCenterBottom
.AddItem FlexPictureAlignmentRightTop & " - RightTop"
.ItemData(.NewIndex) = FlexPictureAlignmentRightTop
.AddItem FlexPictureAlignmentRightCenter & " - RightCenter"
.ItemData(.NewIndex) = FlexPictureAlignmentRightCenter
.AddItem FlexPictureAlignmentRightBottom & " - RightBottom"
.ItemData(.NewIndex) = FlexPictureAlignmentRightBottom
.AddItem FlexPictureAlignmentStretch & " - Stretch"
.ItemData(.NewIndex) = FlexPictureAlignmentStretch
.AddItem FlexPictureAlignmentTile & " - Tile"
.ItemData(.NewIndex) = FlexPictureAlignmentTile
.AddItem FlexPictureAlignmentLeftTopNoOverlap & " - LeftTopNoOverlap"
.ItemData(.NewIndex) = FlexPictureAlignmentLeftTopNoOverlap
.AddItem FlexPictureAlignmentLeftCenterNoOverlap & " - LeftCenterNoOverlap"
.ItemData(.NewIndex) = FlexPictureAlignmentLeftCenterNoOverlap
.AddItem FlexPictureAlignmentLeftBottomNoOverlap & " - LeftBottomNoOverlap"
.ItemData(.NewIndex) = FlexPictureAlignmentLeftBottomNoOverlap
.AddItem FlexPictureAlignmentRightTopNoOverlap & " - RightTopNoOverlap"
.ItemData(.NewIndex) = FlexPictureAlignmentRightTopNoOverlap
.AddItem FlexPictureAlignmentRightCenterNoOverlap & " - RightCenterNoOverlap"
.ItemData(.NewIndex) = FlexPictureAlignmentRightCenterNoOverlap
.AddItem FlexPictureAlignmentRightBottomNoOverlap & " - RightBottomNoOverlap"
.ItemData(.NewIndex) = FlexPictureAlignmentRightBottomNoOverlap
.ListIndex = 0
End With
With Combo2
.AddItem "Generic"
.AddItem "Numeric"
.AddItem "StringNoCase"
.AddItem "String"
.AddItem "Currency"
.AddItem "Date"
.ListIndex = 0
End With
End Sub

Private Sub Form_Resize()
Dim Width As Single, Height As Single
Width = Me.ScaleWidth - VBFlexGrid1.Left - Me.ScaleX(8, vbPixels, Me.ScaleMode)
Height = Me.ScaleHeight - (PicturePanel.Height) - Me.ScaleY(8, vbPixels, Me.ScaleMode)
If Width > 0 Then VBFlexGrid1.Width = Width
If Height > 0 Then VBFlexGrid1.Height = Height
End Sub

Private Sub VBFlexGrid1_ScrollTip(ByVal Row As Long, ByVal Col As Long)
If Row > -1 Then
    VBFlexGrid1.ScrollTipText = "Row " & VBFlexGrid1.TextMatrix(Row, 0)
ElseIf Col > -1 Then
    VBFlexGrid1.ScrollTipText = "Column " & VBFlexGrid1.TextMatrix(0, Col)
End If
End Sub

Private Sub VBFlexGrid1_DividerDblClick(ByVal Row As Long, ByVal Col As Long)
If Row = -1 Then
    VBFlexGrid1.AutoSize Col, , FlexAutoSizeModeColWidth, , , , CBool(VBFlexGrid1.ClipMode = FlexClipModeExcludeHidden)
ElseIf Col = -1 Then
    VBFlexGrid1.AutoSize Row, , FlexAutoSizeModeRowHeight, , , , CBool(VBFlexGrid1.ClipMode = FlexClipModeExcludeHidden)
End If
End Sub

Private Sub Command1_Click()
Set VBFlexGrid1.CellPicture = Picture1.Picture
VBFlexGrid1.CellPictureAlignment = CLng(Combo1.ItemData(Combo1.ListIndex))
End Sub

Private Sub Command2_Click()
Set VBFlexGrid1.CellPicture = Nothing
VBFlexGrid1.CellPictureAlignment = FlexPictureAlignmentLeftTop
End Sub

Private Sub Command3_Click()
If InIDE() = False Then MsgBox "OleCreatePropertyFrame works only in IDE or with an OCX, but not with a compiled UserControl.", vbExclamation + vbOKOnly: Exit Sub
Dim CLSID As OLEGuids.OLECLSID, OldColor As OLE_COLOR
CLSIDFromString StrPtr(CLSID_StandardColorPage), CLSID
PropCellBackColor = VBFlexGrid1.CellBackColor
PropCellForeColor = VBFlexGrid1.CellForeColor
OldColor = PropCellBackColor
OleCreatePropertyFrame Me.hWnd, 0, 0, StrPtr("VBFlexGrid"), 1, Me, 1, CLSID, 0, 0, NULL_PTR
If PropCellBackColor <> OldColor Then VBFlexGrid1.CellBackColor = PropCellBackColor
End Sub

Private Sub Command4_Click()
If InIDE() = False Then MsgBox "OleCreatePropertyFrame works only in IDE or with an OCX, but not with a compiled UserControl.", vbExclamation + vbOKOnly: Exit Sub
Dim CLSID As OLEGuids.OLECLSID, OldColor As OLE_COLOR
CLSIDFromString StrPtr(CLSID_StandardColorPage), CLSID
PropCellBackColor = VBFlexGrid1.CellBackColor
PropCellForeColor = VBFlexGrid1.CellForeColor
OldColor = PropCellForeColor
OleCreatePropertyFrame Me.hWnd, 0, 0, StrPtr("VBFlexGrid"), 1, Me, 1, CLSID, 0, 0, NULL_PTR
If PropCellForeColor <> OldColor Then VBFlexGrid1.CellForeColor = PropCellForeColor
End Sub

Private Sub Command13_Click()
With New InputForm
.Prompt = "ToolTipText for Cell R" & VBFlexGrid1.Row & "C" & VBFlexGrid1.Col
.DefaultText = VBFlexGrid1.CellToolTipText
.Show vbModal, Me
If StrPtr(.Result) <> NULL_PTR Then VBFlexGrid1.CellToolTipText = .Result
End With
End Sub

Private Sub Command5_Click()
With New InputForm
.Prompt = "Text for Cell R" & VBFlexGrid1.Row & "C" & VBFlexGrid1.Col
.DefaultText = VBFlexGrid1.Text
.Show vbModal, Me
If StrPtr(.Result) <> NULL_PTR Then VBFlexGrid1.Text = .Result
End With
End Sub

Private Sub Command10_Click()
VBFlexGrid1.CellEnsureVisible FlexVisibilityCompleteOnly
End Sub

Private Sub Command14_Click()
If InIDE() = False Then MsgBox "OleCreatePropertyFrame works only in IDE or with an OCX, but not with a compiled UserControl.", vbExclamation + vbOKOnly: Exit Sub
Dim CLSID As OLEGuids.OLECLSID, OldFont As StdFont
CLSIDFromString StrPtr(CLSID_StandardFontPage), CLSID
Set PropCellFont = New StdFont
With PropCellFont
.Name = VBFlexGrid1.CellFontName
.Size = VBFlexGrid1.CellFontSize
.Bold = VBFlexGrid1.CellFontBold
.Italic = VBFlexGrid1.CellFontItalic
.Strikethrough = VBFlexGrid1.CellFontStrikeThrough
.Underline = VBFlexGrid1.CellFontUnderline
.Charset = VBFlexGrid1.CellFontCharset
Set OldFont = CloneOLEFont(PropCellFont)
OleCreatePropertyFrame Me.hWnd, 0, 0, StrPtr("VBFlexGrid"), 1, Me, 1, CLSID, 0, 0, NULL_PTR
If .Name <> OldFont.Name Or .Size <> OldFont.Size Or _
.Bold <> OldFont.Bold Or .Italic <> OldFont.Italic Or _
.Strikethrough <> OldFont.Strikethrough Or .Underline <> OldFont.Underline Or _
.Charset <> OldFont.Charset Then
    Set Me.CellFont = PropCellFont
End If
End With
End Sub

Private Sub Command11_Click()
VBFlexGrid1.Clear FlexClearSelection, FlexClearText
End Sub

Private Sub Command12_Click()
VBFlexGrid1.Clear FlexClearSelection, FlexClearEverything
End Sub

Private Sub Command15_Click()
Set Me.CellFont = Nothing
End Sub

Private Sub Command17_Click()
Dim Row1 As Long, Row2 As Long
Dim Col1 As Long, Col2 As Long
With VBFlexGrid1
.GetSelRange Row1, Col1, Row2, Col2
.RowID(.Row) = 1 ' Temporary identification
.Sort = VBA.Choose(Combo2.ListIndex + 1, FlexSortGenericAscending, FlexSortNumericAscending, FlexSortStringNoCaseAscending, FlexSortStringAscending, FlexSortCurrencyAscending, FlexSortDateAscending)
.Row = .RowIndex(1)
.RowID(.RowIndex(1)) = 0 ' Remove temporary identification
.CellEnsureVisible
If Row1 <> Row2 Then .RowSel = IIf(Row1 < .Row, Row1, Row2)
If Col1 <> Col2 Then .ColSel = IIf(Col1 < .Col, Col1, Col2)
End With
End Sub

Private Sub Command18_Click()
Dim Row1 As Long, Row2 As Long
Dim Col1 As Long, Col2 As Long
With VBFlexGrid1
.GetSelRange Row1, Col1, Row2, Col2
.RowID(.Row) = 1 ' Temporary identification
.Sort = VBA.Choose(Combo2.ListIndex + 1, FlexSortGenericDescending, FlexSortNumericDescending, FlexSortStringNoCaseDescending, FlexSortStringDescending, FlexSortCurrencyDescending, FlexSortDateDescending)
.Row = .RowIndex(1)
.RowID(.RowIndex(1)) = 0 ' Remove temporary identification
.CellEnsureVisible
If Row1 <> Row2 Then .RowSel = IIf(Row1 < .Row, Row1, Row2)
If Col1 <> Col2 Then .ColSel = IIf(Col1 < .Col, Col1, Col2)
End With
End Sub

Private Sub Command6_Click()
If InIDE() = False Then MsgBox "OleCreatePropertyFrame works only in IDE or with an OCX, but not with a compiled UserControl.", vbExclamation + vbOKOnly: Exit Sub
Dim SpecifyPages As OLEGuids.ISpecifyPropertyPages, Pages As OLEGuids.OLECAUUID
Set SpecifyPages = VBFlexGrid1.Object
SpecifyPages.GetPages Pages
OleCreatePropertyFrame Me.hWnd, 0, 0, StrPtr("VBFlexGrid"), 1, VBFlexGrid1.Object, Pages.cElems, ByVal Pages.pElems, 0, 0, NULL_PTR
CoTaskMemFree Pages.pElems
Me.SetFocus
End Sub

Private Sub Command7_Click()
Clipboard.Clear
Clipboard.SetData VBFlexGrid1.Picture, IIf(VBFlexGrid1.PictureType <> FlexPictureTypeEnhMetafile, vbCFBitmap, vbCFEMetafile)
MsgBox "You can now paste this printscreen with Ctrl+V in MS Paint for example.", vbInformation + vbOKOnly
End Sub

Private Sub Command19_Click()
UserEditingForm.Show vbModal
End Sub

Private Sub Command8_Click()
VBFlexGrid1.Copy
End Sub

Private Sub Command9_Click()
VBFlexGrid1.Paste
End Sub

Private Sub Picture2_Paint()
Picture2.Cls
Picture2.Print "Drag/drop me"
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PropInsertRowDragging = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PropInsertRowDragging = True Then
    VBFlexGrid1.InsertMarkMode = FlexDropTargetModeByRow
    Picture2.OLEDrag
End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PropInsertRowDragging = False
End Sub

Private Sub Picture2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(Picture2.Name), vbCFRTF
AllowedEffects = vbDropEffectMove + vbDropEffectCopy
End Sub

Private Sub Picture2_OLECompleteDrag(Effect As Long)
PropInsertRowDragging = False
End Sub

Private Sub Command16_Click()
With New InputForm
.SearchMode = True
.Prompt = "Search for cell in scrollable area within column '" & VBFlexGrid1.TextMatrix(0, VBFlexGrid1.Col) & "' (Col = " & VBFlexGrid1.Col & ")"
.Show vbModal, Me
If StrPtr(.Result) <> NULL_PTR Then
    Dim Row As Long
    Row = VBFlexGrid1.FindItem(.Result, , VBFlexGrid1.Col, CBool(Check1.Value = vbChecked))
    If Row > -1 Then
        VBFlexGrid1.Row = Row
        VBFlexGrid1.CellEnsureVisible
        VBFlexGrid1.CellBackColor = vbGreen
    Else
        MsgBox "'" & .Result & "' cannot be found.", vbInformation + vbOKOnly
    End If
End If
End With
End Sub

Private Sub Command20_Click()
VBFlexGrid1.RowHidden(VBFlexGrid1.Row) = Not VBFlexGrid1.RowHidden(VBFlexGrid1.Row)
End Sub

Private Sub Command21_Click()
PropDragRowColActive = Not PropDragRowColActive
MsgBox "DragRowCol mode is '" & PropDragRowColActive & "'." & IIf(PropDragRowColActive, vbLf & "Start dragging while holding down the ALT key.", "")
End Sub

Private Sub VBFlexGrid1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Data.SetData StrToVar(VBFlexGrid1.Name), vbCFRTF
AllowedEffects = vbDropEffectMove
End Sub

Private Sub VBFlexGrid1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Data.GetFormat(vbCFRTF) = False Then
    Effect = vbDropEffectNone
    Exit Sub
End If
Dim DataString As String
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = VBFlexGrid1.Name Or DataString = Picture2.Name Then
    If PropInsertRowDragging = True Then
        Effect = vbDropEffectCopy Or vbDropEffectMove
    Else
        Effect = vbDropEffectNone
        Exit Sub
    End If
Else
    Effect = vbDropEffectNone
    Exit Sub
End If
If State = vbOver Then
    If PropInsertRowDragging = True Then
        Dim HitInsertMark As Long, After As Boolean
        With VBFlexGrid1
        HitInsertMark = .HitTestInsertMark(X, Y, After)
        If HitInsertMark >= .FixedRows Then
            .InsertMark(After) = HitInsertMark
        Else
            .InsertMark = .FixedRows
        End If
        End With
    End If
ElseIf State = vbLeave Then
    If PropInsertRowDragging = True Then VBFlexGrid1.InsertMark = -1
End If
End Sub

Private Sub VBFlexGrid1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(vbCFRTF) = False Then Exit Sub
Dim DataString As String
DataString = VarToStr(Data.GetData(vbCFRTF))
If DataString = VBFlexGrid1.Name Or DataString = Picture2.Name Then
    If PropInsertRowDragging = True Then
        Dim HitInsertMark As Long, After As Boolean
        With VBFlexGrid1
        HitInsertMark = .HitTestInsertMark(X, Y, After)
        If HitInsertMark >= .FixedRows Then
            .AddItem CStr(.Rows - .FixedRows + 1) & vbTab & "New", IIf(After = True, HitInsertMark + 1, HitInsertMark)
        Else
            .AddItem CStr(.Rows - .FixedRows + 1) & vbTab & "New", .FixedRows
        End If
        .InsertMark = -1
        End With
        PropInsertRowDragging = False
    End If
End If
End Sub

Private Sub VBFlexGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
With VBFlexGrid1
If PropDragRowColActive = True And Button = vbLeftButton And (Shift And vbAltMask) = vbAltMask Then
    .HitTest X, Y
    If .HitResult = FlexHitResultCell Then
        If .HitCol < .FixedCols And .HitRow >= .FixedRows Then
            VBFlexGrid1.RowPosition(.HitRow) = VBFlexGrid1.DragRow(.HitRow)
            Cancel = True
        ElseIf .HitCol >= .FixedCols And .HitRow < .FixedRows Then
            VBFlexGrid1.ColPosition(.HitCol) = VBFlexGrid1.DragCol(.HitCol)
            Cancel = True
        End If
    End If
End If
End With
End Sub
